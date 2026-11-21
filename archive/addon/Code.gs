/** @OnlyCurrentDoc */
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Open phooey! finance", "openSidebar")
    .addToUi();
}
function onInstall() {
  onOpen();
}
function openSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("phooey! finance"),
  );
}
function book_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}
function props_() {
  return PropertiesService.getUserProperties();
}
function key_() {
  return "workbook:" + book_().getId();
}
function api_(path, data, connection) {
  const grant =
    connection || JSON.parse(props_().getProperty(key_()) || "null");
  if (!grant) throw Error("Pair your workbook with a code from Phooey.");
  const response = UrlFetchApp.fetch(grant.origin + "/addon/" + path, {
    method: "post",
    followRedirects: false,
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + (grant.token || ""),
      "X-Phooey-Workbook": book_().getId(),
    },
    payload: JSON.stringify(data || {}),
  });
  let result;
  try {
    result = JSON.parse(response.getContentText());
  } catch (_) {
    throw Error("Your workspace is unavailable.");
  }
  if (response.getResponseCode() !== 200)
    throw Error(result.error || "Request failed.");
  return result;
}
function pairWorkbook(origin, code) {
  const suffix =
    PropertiesService.getScriptProperties().getProperty("WORKSPACE_DOMAIN");
  if (
    !suffix ||
    !/^https:\/\/[a-z0-9-]+\.[a-z0-9.-]+$/.test(origin) ||
    !origin.endsWith("." + suffix)
  )
    throw Error("Use the workspace address shown in Phooey.");
  const grant = api_("pair", { code: code.trim() }, { origin: origin });
  props_().setProperty(
    key_(),
    JSON.stringify({ origin: origin, token: grant.token }),
  );
  return "Workbook paired.";
}
function refreshFeed() {
  api_("refresh");
  return "Workbook refreshed.";
}
function sheet_(number) {
  const sheet = book_()
    .getSheets()
    .find((s) => s.getSheetId() === number);
  if (!sheet) throw Error("Use the workbook created by Phooey.");
  const columns = [
    "ID",
    "Revision",
    "Date",
    "Account",
    "Description",
    "Amount",
    "Currency",
    "Category",
    "Type",
    "Merchant",
    "Notes",
    "Tags",
    "Reviewed",
    "Pending",
  ];
  if (number === 105) columns.push("Request key");
  const header = sheet.getRange(1, 1, 1, columns.length).getValues()[0];
  if (columns.some((label, i) => header[i] !== label))
    throw Error(
      "Restore the original column order before editing. Keep custom columns to the right.",
    );
  return sheet;
}
function asText_(value) {
  return typeof value === "string" && /^[=+@-]/.test(value)
    ? "'" + value
    : value;
}
function selected_() {
  const sheet = book_().getActiveSheet(),
    range = sheet.getActiveRange();
  if (![101, 105].includes(sheet.getSheetId()) || range.getRow() < 2)
    throw Error("Select a transaction or a draft row.");
  sheet_(sheet.getSheetId());
  return {
    sheet: sheet,
    row: range.getRow(),
    values: sheet
      .getRange(range.getRow(), 1, 1, sheet.getSheetId() === 105 ? 15 : 14)
      .getValues()[0],
  };
}
function stageRow(blank) {
  const sheet = sheet_(105);
  const values = blank
    ? [
        "",
        "",
        Utilities.formatDate(
          new Date(),
          book_().getSpreadsheetTimeZone(),
          "yyyy-MM-dd",
        ),
        "",
        "",
        "",
        "USD",
        "",
        "expense",
        "",
        "",
        '["Business"]',
        true,
        false,
      ]
    : selected_().values.slice(0, 14);
  values.push(Utilities.getUuid());
  const row = sheet.getLastRow() + 1;
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  sheet.getRange(row, 1, 1, 15).setValues([values.map(asText_)]);
  sheet.activate();
  sheet.setActiveRange(sheet.getRange(row, 4));
  return "Edit the draft cells, then choose Save selected drafts.";
}
function selectId_(list, value, label) {
  const hits = list.filter((i) => i.id === value || i.name === value);
  if (hits.length !== 1)
    throw Error("Enter a unique " + label + " name or ID.");
  return hits[0];
}
function money_(value) {
  if (value === "" || !Number.isFinite(Number(value)))
    throw Error("Enter a numeric amount.");
  const amount = Number(value) * 100;
  if (Math.abs(amount - Math.round(amount)) > 0.000001)
    throw Error("Use at most two decimal places.");
  return Math.round(amount);
}
function bool_(v) {
  if (v === true || v === false) return v;
  if (String(v).toLowerCase() === "true") return true;
  if (String(v).toLowerCase() === "false" || v === "") return false;
  throw Error("Use TRUE or FALSE.");
}
function draft_(sel, opts, splits) {
  const v = sel.values,
    account = selectId_(opts.accounts, String(v[3]), "account");
  if (account.currency !== String(v[6]))
    throw Error("Currency must match the selected account.");
  const date =
    v[2] instanceof Date
      ? Utilities.formatDate(
          v[2],
          book_().getSpreadsheetTimeZone(),
          "yyyy-MM-dd",
        )
      : String(v[2]);
  let tags;
  try {
    tags = JSON.parse(String(v[11]) || "[]");
  } catch (_) {
    throw Error('Tags must be a list, for example ["Business"].');
  }
  return {
    account_id: account.id,
    posted: date,
    description: String(v[4]),
    amount: money_(v[5]),
    category_id: v[7]
      ? selectId_(opts.categories, String(v[7]), "category").id
      : null,
    type: String(v[8]),
    merchant: String(v[9]),
    notes: String(v[10]),
    tags: tags,
    reviewed: bool_(v[12]),
    pending: bool_(v[13]),
    splits: splits,
  };
}
function saveDraft() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sel = selected_();
    if (sel.sheet.getSheetId() !== 105)
      throw Error("Stage a draft before saving.");
    const count = sel.sheet.getActiveRange().getNumRows();
    if (count > 50) throw Error("Select at most 50 draft rows per save.");
    const options = api_("options");
    for (let offset = 0; offset < count; offset++) {
      const row = sel.row + offset;
      const values = sel.sheet.getRange(row, 1, 1, 15).getValues()[0];
      try {
        saveOneDraft_({ row: row, sheet: sel.sheet, values: values }, options);
      } catch (e) {
        throw Error(
          "Stopped at row " +
            row +
            ". " +
            offset +
            " earlier rows saved. " +
            e.message,
        );
      }
    }
    return count + " draft(s) saved. Refresh the feed to update reports.";
  } finally {
    lock.releaseLock();
  }
}
function saveOneDraft_(sel, options) {
  const v = sel.values,
    prior = v[0] ? api_("transaction", { id: v[0] }) : null;
  if (!v[14]) {
    v[14] = Utilities.getUuid();
    sel.sheet.getRange(sel.row, 15).setValue(v[14]);
    SpreadsheetApp.flush();
  }
  // Persist the request key before the write, so a lost response can be safely retried.
  const saved = api_("save", {
    id: v[0] || undefined,
    revision: v[0] ? Number(v[1]) : undefined,
    requestKey: v[14],
    transaction: draft_(sel, options, prior ? prior.splits : []),
  });
  sel.sheet.getRange(sel.row, 1, 1, 2).setValues([[saved.id, saved.revision]]);
  sel.sheet.getRange(sel.row, 15).setValue(Utilities.getUuid());
}
function reloadDraft(id) {
  const sel = selected_();
  if (sel.sheet.getSheetId() !== 105 || sel.values[0] !== id)
    throw Error("Select the reviewed draft.");
  const t = api_("transaction", { id: id });
  const values = [
    t.id,
    t.revision,
    t.posted,
    t.account_name,
    t.description,
    t.amount / 100,
    t.currency,
    t.category_name || "",
    t.type,
    t.merchant,
    t.notes,
    JSON.stringify(t.tags),
    !!t.reviewed,
    !!t.pending,
    Utilities.getUuid(),
  ];
  sel.sheet.getRange(sel.row, 1, 1, 15).setValues([values.map(asText_)]);
  return "The draft now matches the current ledger.";
}
function inspectRow() {
  const sel = selected_(),
    id = sel.values[0];
  if (!id) throw Error("Save this transaction first.");
  const current = api_("transaction", { id: id }),
    options = api_("options");
  return {
    id: id,
    revision: current.revision,
    draftRevision: sel.values[1],
    description: current.description,
    amount: current.amount,
    notes: current.notes,
    category: current.category_name,
    posted: current.posted,
    account: current.account_name,
    merchant: current.merchant,
    tags: current.tags,
    splits: current.splits,
    transfer: current.transfer_id,
    categories: options.categories.map((c) => ({ id: c.id, name: c.name })),
  };
}
function acceptRevision(id, revision) {
  const sel = selected_();
  if (sel.sheet.getSheetId() !== 105 || sel.values[0] !== id)
    throw Error("Select the reviewed draft.");
  sel.sheet.getRange(sel.row, 2).setValue(revision);
  sel.sheet.getRange(sel.row, 15).setValue(Utilities.getUuid());
  return "Draft kept. Review the cells, then save.";
}
function updateSplits(id, revision, splits) {
  const tx = api_("transaction", { id: id });
  if (tx.revision !== revision)
    throw Error("Review the latest transaction first.");
  const data = {
    account_id: tx.account_id,
    posted: tx.posted,
    amount: tx.amount,
    description: tx.description,
    merchant: tx.merchant,
    category_id: tx.category_id,
    type: tx.type,
    notes: tx.notes,
    tags: tx.tags,
    reviewed: !!tx.reviewed,
    pending: !!tx.pending,
    splits: splits.map((s) => ({
      amount: money_(s.amount),
      type: s.type,
      category_id: s.category || null,
      notes: s.notes || "",
    })),
  };
  api_("save", {
    id: id,
    revision: revision,
    requestKey: Utilities.getUuid(),
    transaction: data,
  });
  return "Split saved. Refresh your feed.";
}
function deleteReviewed(id, revision) {
  api_("delete", { id: id, revision: revision });
  return "Transaction deleted. Refresh your feed.";
}
function unlinkReviewed(id, revision) {
  api_("unlink", { id: id, revision: revision });
  return "Transfer unlinked.";
}
function linkSelected() {
  const sel = selected_(),
    a = sel.sheet.getActiveRange();
  if (a.getNumRows() !== 2)
    throw Error("Select two adjacent transaction rows to link.");
  const rows = sel.sheet.getRange(a.getRow(), 1, 2, 2).getValues();
  api_("link", {
    first: rows[0][0],
    second: rows[1][0],
    revisions: [rows[0][1], rows[1][1]],
  });
  return "Transfer linked. Refresh your feed.";
}
