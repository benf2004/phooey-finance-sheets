// Install in a new PRIVATE standalone Apps Script project. See the setup page.
// Only the three Tiller feed tabs and the connector's account map receive writes.
const MAP_TAB = "Phooey Account Map";
const MAP_HEADER = [
  "Phooey ID",
  "Phooey account",
  "Currency",
  "Tiller Account ID or NEW",
];
const PREFIX = "phooey:";
function settings_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const file = (p.WORKBOOK_URL || "").match(
    /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([\w-]+)(?:\/|$)/,
  );
  if (
    !file ||
    !/^https:\/\/[a-z0-9.-]+(?::\d+)?$/.test(p.PHOOEY_ORIGIN || "") ||
    !p.READ_TOKEN
  )
    throw Error(
      "Set WORKBOOK_URL, PHOOEY_ORIGIN and READ_TOKEN in this private project's script properties.",
    );
  if (
    !/^\d{4}-\d\d-\d\d$/.test(p.FIRST_SYNC_DATE || "") ||
    new Date(p.FIRST_SYNC_DATE + "T12:00:00Z").toISOString().slice(0, 10) !==
      p.FIRST_SYNC_DATE
  )
    throw Error(
      "Set FIRST_SYNC_DATE to the first day after your last Tiller fill (YYYY-MM-DD).",
    );
  return {
    file: file[1],
    origin: p.PHOOEY_ORIGIN,
    token: p.READ_TOKEN,
    from: p.FIRST_SYNC_DATE,
  };
}
function rpc_(config, method, params) {
  const r = UrlFetchApp.fetch(config.origin + "/mcp", {
    method: "post",
    contentType: "application/json",
    followRedirects: false,
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + config.token,
      Accept: "application/json, text/event-stream",
      "MCP-Protocol-Version": "2025-03-26",
    },
    payload: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: method,
      params: params,
    }),
  });
  if (r.getResponseCode() !== 200)
    throw Error(
      "Phooey returned " +
        r.getResponseCode() +
        ". Check the workspace address, subscription, and read token.",
    );
  const text = r.getContentText().trim();
  const msg = JSON.parse(
    text.startsWith("{")
      ? text
      : text
          .split(/\r?\n/)
          .find((s) => s.startsWith("data:"))
          .slice(5),
  );
  if (msg.error || !msg.result)
    throw Error("Phooey could not complete the request.");
  return msg.result;
}
function tool_(config, name, args) {
  const r = rpc_(config, "tools/call", { name: name, arguments: args || {} });
  if (r.isError)
    throw Error(name + " failed. " + String(r.content[0].text).slice(0, 160));
  return JSON.parse(r.content[0].text);
}
function source_(c) {
  if (rpc_(c, "tools/list", {}).tools.some((t) => t.name === "get_tiller_feed"))
    return tool_(c, "get_tiller_feed", { from: c.from });
  // Older self-hosted versions expose the same ledger through paginated read tools.
  const accounts = tool_(c, "list_accounts"),
    categories = tool_(c, "list_categories"),
    transactions = [],
    seen = new Set();
  let total;
  for (let page = 1; ; page++) {
    const result = tool_(c, "search_transactions", {
      from: c.from,
      pending: "0",
      page: page,
      limit: 100,
    });
    if (result.total > 5000 || (total !== undefined && total !== result.total))
      throw Error(
        "Ledger changed or exceeds 5,000 transactions. Retry with a narrower date range.",
      );
    total = result.total;
    for (const t of result.items) {
      if (seen.has(t.id))
        throw Error("Ledger changed during the read. Retry sync.");
      seen.add(t.id);
      const detail = t.split_count
        ? tool_(c, "get_transaction", { id: t.id })
        : null;
      if (detail && detail.revision !== t.revision)
        throw Error("A split changed. Retry sync.");
      (detail ? detail.splits : [t]).forEach((s, i) =>
        transactions.push({
          id: t.id,
          part: i,
          revision: t.revision,
          posted: t.posted,
          description: t.description,
          tags: t.tags,
          account_id: t.account_id,
          category_id: s.category_id,
          type: s.type,
          amount: s.amount,
        }),
      );
    }
    if (seen.size === total) break;
    if (!result.items.length)
      throw Error("Incomplete transaction page. Retry sync.");
  }
  return {
    accounts: accounts,
    categories: categories,
    transactions: transactions,
    balances: [],
  };
}
function table_(book, name, required) {
  const sheet = book.getSheetByName(name);
  if (!sheet)
    throw Error(
      "The selected workbook is missing " +
        name +
        ". Use a native Google Sheets copy of the Tiller workbook.",
    );
  const range = sheet.getDataRange(),
    values = range.getValues(),
    formulas = range.getFormulas();
  const header = sheet
    .getRange(1, 1, 1, range.getNumColumns())
    .getDisplayValues()[0]
    .map((s) => s.trim());
  for (const h of required)
    if (header.filter((v) => v === h).length !== 1)
      throw Error(name + " needs a unique " + h + " column.");
  return {
    name: name,
    id: sheet.getSheetId(),
    header: header,
    values: values,
    formulas: formulas,
    tz: book.getSpreadsheetTimeZone(),
    capacity: sheet.getMaxRows(),
    next: values.length,
  };
}
function readTables_(book) {
  return {
    tx: table_(book, "Transactions", [
      "Date",
      "Description",
      "Category",
      "Amount",
      "Account",
      "Transaction ID",
      "Account ID",
    ]),
    hist: table_(book, "Balance History", [
      "Date",
      "Account",
      "Account ID",
      "Balance ID",
      "Balance",
    ]),
    cats: table_(book, "Categories", ["Category", "Group", "Type"]),
    rules: book.getSheetByName("AutoCat")
      ? table_(book, "AutoCat", ["Category"])
      : null,
  };
}
function lookup_(t, row, field) {
  const i = t.header.indexOf(field);
  return i < 0 ? "" : (row[i] ?? "");
}
function serial_(d) {
  const n = Date.parse(d + "T00:00:00Z");
  if (!Number.isFinite(n)) throw Error("Invalid posted date.");
  return n / 86400000 + 25569;
}
function dates_(d) {
  const n = serial_(d);
  return {
    Date: n,
    Month: serial_(d.slice(0, 7) + "-01"),
    Week: n - new Date(d + "T00:00:00Z").getUTCDay(),
  };
}
function day_(value, tz) {
  return value instanceof Date
    ? Utilities.formatDate(value, tz, "yyyy-MM-dd")
    : typeof value === "number"
      ? new Date((Math.floor(value) - 25569) * 86400000)
          .toISOString()
          .slice(0, 10)
      : String(value || "").slice(0, 10);
}
function value_(v) {
  if (v === "" || v == null) return {};
  return {
    userEnteredValue:
      typeof v === "number" ? { numberValue: v } : { stringValue: String(v) },
  };
}
function sameValue_(value, desired, key, tz) {
  if (value instanceof Date) {
    if (key === "Time") {
      const [hours, minutes, seconds] = Utilities.formatDate(
        value,
        tz,
        "HH:mm:ss",
      )
        .split(":")
        .map(Number);
      return (
        Math.abs((hours * 3600 + minutes * 60 + seconds) / 86400 - desired) <
        1 / 86400
      );
    }
    return serial_(day_(value, tz)) === desired;
  }
  return (value ?? "") === (desired ?? "");
}
function writeRow_(requests, t, row, fields) {
  // Use typed API values. Descriptions beginning with '=' remain literal text.
  // Update contiguous managed columns only. Existing formulas are always preserved.
  let start = -1,
    cells = [];
  function flush() {
    if (cells.length)
      requests.push({
        updateCells: {
          start: { sheetId: t.id, rowIndex: row, columnIndex: start },
          rows: [{ values: cells }],
          fields: "userEnteredValue",
        },
      });
    cells = [];
    start = -1;
  }
  t.header.forEach((key, col) => {
    if (
      !Object.prototype.hasOwnProperty.call(fields, key) ||
      t.formulas[0]?.[col] ||
      t.formulas[row]?.[col] ||
      sameValue_(t.values[row]?.[col], fields[key], key, t.tz)
    ) {
      flush();
      return;
    }
    if (start < 0) start = col;
    cells.push(value_(fields[key]));
  });
  flush();
}
function idIndex_(t, label) {
  const map = new Map();
  t.values.slice(1).forEach((r, i) => {
    const id = String(lookup_(t, r, label));
    if (id.startsWith(PREFIX)) {
      if (map.has(id))
        throw Error(
          "Duplicate Phooey ID in " + t.name + ". Resolve it before syncing.",
        );
      map.set(id, i + 1);
    }
  });
  return map;
}
function mappings_(mapping, accounts, hist, tz) {
  const result = new Map(),
    used = new Set(),
    known = new Map();
  for (const r of hist.values.slice(1)) {
    const key = String(lookup_(hist, r, "Account ID")),
      date = day_(lookup_(hist, r, "Date"), tz);
    if (key && (!known.has(key) || date >= known.get(key).date))
      known.set(key, {
        date: date,
        account: lookup_(hist, r, "Account"),
        number: lookup_(hist, r, "Account #"),
        institution: lookup_(hist, r, "Institution"),
        type: lookup_(hist, r, "Type"),
        cls: lookup_(hist, r, "Class"),
      });
  }
  for (const r of mapping.values.slice(1)) {
    const choice = String(r[3] || "").trim();
    if (!choice) continue;
    const account = accounts.find((a) => a.id === r[0]);
    if (!account) continue;
    if (account.currency !== "USD")
      throw Error(
        "The US tax workbook accepts USD accounts. Leave other currencies unmapped.",
      );
    const key = choice === "NEW" ? PREFIX + "account:" + account.id : choice;
    if (used.has(key) || result.has(account.id))
      throw Error("Each Phooey and Tiller account must be mapped once.");
    if (choice !== "NEW" && !known.has(key))
      throw Error(
        "A mapped Tiller account ID is missing from Balance History.",
      );
    const existing = known.get(key) || {};
    result.set(account.id, {
      key: key,
      account: existing.account || account.name,
      number: existing.number || "",
      institution: existing.institution || account.institution,
      type: existing.type || account.kind,
      cls:
        existing.cls ||
        (["loan", "credit"].includes(account.kind) ? "Liability" : "Asset"),
    });
    used.add(key);
  }
  if (!result.size)
    throw Error("Choose at least one account in Phooey Account Map.");
  return result;
}
function autoCategory_(rules, tx, account) {
  if (!rules) return "";
  for (const r of rules.values.slice(1)) {
    let matched = true,
      hasFilter = false;
    const cat = lookup_(rules, r, "Category");
    if (!cat) continue;
    for (let i = 0; i < r.length; i++) {
      const k = rules.header[i],
        v = r[i];
      if (v === "" || v == null || k === "Category") continue;
      hasFilter = true;
      const vals = {
        "Description Contains": tx.description,
        "Account Contains": account.account,
        "Institution Contains": account.institution,
      };
      if (k in vals)
        matched =
          matched &&
          String(vals[k]).toLowerCase().includes(String(v).toLowerCase());
      else if (k === "Amount Min")
        matched = matched && tx.amount / 100 >= Number(v);
      else if (k === "Amount Max")
        matched = matched && tx.amount / 100 <= Number(v);
      else matched = false;
    }
    if (hasFilter && matched) return String(cat);
  }
  return "";
}
function plan_(tables, mapping, data, from, tz) {
  const { tx, hist, cats, rules } = tables,
    requests = [],
    map = mappings_(mapping, data.accounts, hist, tz);
  const txIds = idIndex_(tx, "Transaction ID"),
    balanceIds = idIndex_(hist, "Balance ID"),
    catNames = new Set(
      cats.values
        .slice(1)
        .map((r) => String(lookup_(cats, r, "Category")).toLowerCase()),
    );
  const mappedKeys = new Set([...map.values()].map((a) => a.key));
  for (const r of tx.values.slice(1)) {
    if (
      !String(lookup_(tx, r, "Transaction ID")).startsWith(PREFIX) &&
      mappedKeys.has(String(lookup_(tx, r, "Account ID"))) &&
      day_(lookup_(tx, r, "Date"), tz) >= from
    )
      throw Error(
        "Existing Tiller transactions overlap the first sync date. Stop the old feed and choose a date after its last transaction.",
      );
  }
  for (const r of hist.values.slice(1)) {
    if (
      !String(lookup_(hist, r, "Balance ID")).startsWith(PREFIX) &&
      mappedKeys.has(String(lookup_(hist, r, "Account ID"))) &&
      day_(lookup_(hist, r, "Date"), tz) >= from
    )
      throw Error(
        "Existing Tiller balances overlap the first sync date. Choose a date after the last Tiller fill.",
      );
  }
  const seen = new Set(),
    details = new Map(data.categories.map((c) => [c.id, c]));
  for (const t of data.transactions) {
    const account = map.get(t.account_id);
    if (!account) continue;
    const key = PREFIX + "tx:" + t.id + ":" + t.part;
    if (seen.has(key)) throw Error("Duplicate source transaction.");
    seen.add(key);
    const c = details.get(t.category_id),
      business =
        (t.tags.some((tag) => String(tag).toLowerCase() === "business") ||
          ["business_income", "business_expense"].includes(c?.tax_treatment)) &&
        t.type !== "transfer" &&
        !(c && ["excluded", "tax_payment"].includes(c.tax_treatment));
    const row = txIds.get(key),
      localCategory =
        row === undefined ? "" : lookup_(tx, tx.values[row], "Category");
    const category =
      c?.name ||
      localCategory ||
      autoCategory_(rules, t, account) ||
      "Phooey " + (business ? "business " : "") + t.type;
    if (!catNames.has(category.toLowerCase())) {
      writeRow_(requests, cats, cats.next++, {
        Category: category,
        Group: business ? "Business" : "Other",
        Type: t.type.charAt(0).toUpperCase() + t.type.slice(1),
        Tags: business ? "Business" : "",
      });
      catNames.add(category.toLowerCase());
    }
    const fields = {
      ...dates_(t.posted),
      Description: t.description,
      "Full Description": t.description,
      Category: category,
      Amount: t.amount / 100,
      Account: account.account,
      "Account #": account.number,
      "Account ID": account.key,
      "Transaction ID": key,
      Institution: account.institution,
    };
    if (row === undefined) fields["Date Added"] = serial_(t.posted);
    writeRow_(requests, tx, row ?? tx.next++, fields);
  }
  // Remove only feed-owned values of previously published allocations which no
  // longer exist in this complete snapshot (re-splits, transfers or deletions).
  for (const [key, row] of txIds)
    if (
      !seen.has(key) &&
      mappedKeys.has(String(lookup_(tx, tx.values[row], "Account ID"))) &&
      day_(lookup_(tx, tx.values[row], "Date"), tz) >= from
    ) {
      writeRow_(
        requests,
        tx,
        row,
        Object.fromEntries(
          [
            "Date",
            "Month",
            "Week",
            "Description",
            "Full Description",
            "Category",
            "Amount",
            "Account",
            "Account #",
            "Account ID",
            "Transaction ID",
            "Institution",
            "Date Added",
          ].map((k) => [k, ""]),
        ),
      );
    }
  const samples = new Map();
  for (const b of [
    ...data.balances,
    ...data.accounts
      .filter((a) => a.balance_at)
      .map((a) => ({
        account_id: a.id,
        balance: a.balance,
        captured_at: a.balance_at,
      })),
  ]) {
    const a = map.get(b.account_id);
    if (!a) continue;
    const time = new Date(
      String(b.captured_at).replace(
        /^(\d{4}-\d\d-\d\d) (\d\d:\d\d:\d\d)$/,
        "$1T$2Z",
      ),
    );
    if (!Number.isFinite(time.getTime())) continue;
    const day = day_(time, tz);
    if (day < from) continue;
    // Accounts looks up the most recent date. Keep its last observation unique even with two fills on the same day.
    const key = PREFIX + "balance:" + b.account_id + ":" + day;
    if (samples.has(key) && samples.get(key).time > time) continue;
    samples.set(key, {
      a: a,
      time: time,
      day: day,
      balance: b.balance,
    });
  }
  for (const [key, b] of samples) {
    const row = balanceIds.get(key),
      time = Utilities.formatDate(b.time, "UTC", "HH:mm:ss")
        .split(":")
        .map(Number);
    writeRow_(requests, hist, row ?? hist.next++, {
      ...dates_(b.day),
      Time: (time[0] * 3600 + time[1] * 60 + time[2]) / 86400,
      Account: b.a.account,
      "Account ID": b.a.key,
      "Account #": b.a.number,
      Institution: b.a.institution,
      "Balance ID": key,
      Balance:
        (String(b.a.cls).toLowerCase() === "liability"
          ? -b.balance
          : b.balance) / 100,
      Class: b.a.cls,
      Type: b.a.type,
      "Account Status": "ACTIVE",
      "Date Added": serial_(b.day),
    });
  }
  const dimensions = [];
  for (const t of [tx, hist, cats]) {
    if (t.next > t.capacity)
      dimensions.push({
        appendDimension: {
          sheetId: t.id,
          dimension: "ROWS",
          length: t.next - t.capacity,
        },
      });
    if (t.next === t.values.length) continue;
    for (const k of ["Date", "Date Added", "Month", "Week", "Time"]) {
      const col = t.header.indexOf(k);
      if (col < 0 || t.formulas[0]?.[col]) continue;
      dimensions.push({
        repeatCell: {
          range: {
            sheetId: t.id,
            startRowIndex: t.values.length,
            endRowIndex: t.next,
            startColumnIndex: col,
            endColumnIndex: col + 1,
          },
          cell: {
            userEnteredFormat: {
              numberFormat:
                k === "Time"
                  ? { type: "TIME", pattern: "h:mm am/pm" }
                  : {
                      type: "DATE",
                      pattern: k === "Month" ? "mmm yyyy" : "yyyy-mm-dd",
                    },
            },
          },
          fields: "userEnteredFormat.numberFormat",
        },
      });
    }
  }
  return {
    requests: dimensions.concat(requests),
    transactions: seen.size,
    balances: samples.size,
  };
}
function prepareWorkbook() {
  const c = settings_(),
    book = SpreadsheetApp.openById(c.file);
  readTables_(book);
  const accounts = tool_(c, "list_accounts");
  const categories = book.getSheetByName("Categories");
  const cat = table_(book, "Categories", ["Category", "Group", "Type"]);
  if (!cat.header.includes("Tags")) {
    const col = cat.header.length + 1;
    if (col > categories.getMaxColumns()) categories.insertColumnAfter(col - 1);
    categories.getRange(1, col).setValue("Tags");
  }
  let sheet = book.getSheetByName(MAP_TAB);
  if (!sheet) {
    sheet = book.insertSheet(MAP_TAB);
    sheet.getRange(1, 1, 1, 4).setValues([MAP_HEADER]);
    sheet.setFrozenRows(1);
  }
  const mapping = table_(book, MAP_TAB, MAP_HEADER);
  if (!MAP_HEADER.every((h, i) => mapping.header[i] === h))
    throw Error("Restore the four account-map headers.");
  const known = new Set(mapping.values.slice(1).map((r) => r[0]));
  const fresh = accounts
    .filter((a) => !known.has(a.id))
    .map((a) => ({ values: [a.id, a.name, a.currency, ""].map(value_) }));
  if (fresh.length)
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            appendCells: {
              sheetId: sheet.getSheetId(),
              rows: fresh,
              fields: "userEnteredValue",
            },
          },
        ],
      },
      c.file,
    );
  console.log(
    "Account map ready. Choose the existing Tiller ID or enter NEW for each new USD account.",
  );
}
function runSync_(write) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const c = settings_(),
      book = SpreadsheetApp.openById(c.file),
      tables = readTables_(book),
      mapping = table_(book, MAP_TAB, MAP_HEADER);
    if (!MAP_HEADER.every((h, i) => mapping.header[i] === h))
      throw Error("Restore the original account map headers.");
    if (!tables.cats.header.includes("Tags"))
      throw Error("Run prepareWorkbook to add the Tags column.");
    const props = PropertiesService.getScriptProperties(),
      signature = JSON.stringify([c.file, c.origin, c.from]);
    if (
      props.getProperty("SYNC_BINDING") &&
      props.getProperty("SYNC_BINDING") !== signature
    )
      throw Error(
        "This project already syncs another workbook, ledger or start date. Use a new private project for a new connection.",
      );
    const previous = JSON.parse(props.getProperty("SYNC_ACCOUNTS") || "{}"),
      choices = {};
    for (const r of mapping.values.slice(1))
      if (r[0] && r[3]) choices[r[0]] = String(r[3]).trim();
    for (const [id, choice] of Object.entries(previous))
      if (choices[id] && choices[id] !== choice)
        throw Error(
          "An existing account mapping changed. Restore its original Tiller account ID before syncing.",
        );
    const result = plan_(
      tables,
      mapping,
      source_(c),
      c.from,
      book.getSpreadsheetTimeZone(),
    );
    if (write && result.requests.length)
      Sheets.Spreadsheets.batchUpdate({ requests: result.requests }, c.file);
    if (write)
      props.setProperties({
        SYNC_BINDING: signature,
        SYNC_ACCOUNTS: JSON.stringify({ ...previous, ...choices }),
      });
    console.log(
      (write ? "Synced " : "Preview: ") +
        result.transactions +
        " posted transaction allocations and " +
        result.balances +
        " balance observations.",
    );
  } finally {
    lock.releaseLock();
  }
}
function previewSync() {
  runSync_(false);
}
function syncWorkbook() {
  runSync_(true);
}
function enableDailySync() {
  if (
    !ScriptApp.getProjectTriggers().some(
      (t) => t.getHandlerFunction() === "syncWorkbook",
    )
  )
    ScriptApp.newTrigger("syncWorkbook")
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
}
function disableDailySync() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "syncWorkbook")
    .forEach((t) => ScriptApp.deleteTrigger(t));
}
