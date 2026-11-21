import { reportTabs, syncedReports } from "./foundation";
import { z } from "zod";
import { decrypt, encrypt } from "../server/auth";
import { Ledger } from "../server/ledger";
import { taxState } from "../server/taxes";
import {
  assert,
  errorMessage,
  hash,
  id,
  now,
  randomToken,
} from "../server/util";
import { challenge, consume, type HostedEnv } from "./control";

const transactionColumns = [
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
export const tabs = [
  { id: 101, name: "Transactions", columns: transactionColumns },
  {
    id: 102,
    name: "Accounts & Balances",
    columns: ["Account", "Type", "Currency", "Balance", "Balance as of", "ID"],
  },
  {
    id: 103,
    name: "Monthly Cash Flow",
    columns: ["Month", "Currency", "Income", "Expenses", "Net"],
  },
  {
    id: 104,
    name: "Quarterly Taxes",
    columns: [
      "Year",
      "Cutoff",
      "Tag",
      "Jurisdiction",
      "Quarter",
      "Due",
      "Estimate (USD)",
      "Shortfall (USD)",
    ],
  },
  { id: 105, name: "Edits", columns: [...transactionColumns, "Request key"] },
  { id: 106, name: "Tax Assumptions", columns: ["Setting", "Value"] },
  ...reportTabs,
];
type Cell = string | number | boolean | null;
type Book = {
  creation_key: string;
  tax_year: number | null;
  file_id: string | null;
  generation: string;
  refresh_secret: string;
  grant_hash: string | null;
  lease_until: number;
  synced_at: number | null;
  error: string | null;
};
const getBook = (env: HostedEnv) =>
  env.DB.prepare("SELECT * FROM hosted_workbook WHERE id=1").first<Book>();
export async function sheetStatus(env: HostedEnv) {
  const b = await getBook(env);
  return {
    year: b?.tax_year ?? new Date().getUTCFullYear(),
    connected: !!b,
    url: b?.file_id
      ? "https://docs.google.com/spreadsheets/d/" + b.file_id + "/edit"
      : null,
    synced: b?.synced_at,
    error: b?.error,
    paired: !!b?.grant_hash,
  };
}
export async function connectGoogle(
  env: HostedEnv,
  customer: string,
  origin: string,
) {
  assert(
    env.GOOGLE_CLIENT && env.GOOGLE_SECRET,
    503,
    "Google Sheets is awaiting configuration.",
  );
  const state = await challenge(env, "google:" + origin, { customer });
  const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  url.search = new URLSearchParams({
    client_id: env.GOOGLE_CLIENT,
    redirect_uri: origin + "/cloud/google/callback",
    response_type: "code",
    scope: "https://www.googleapis.com/auth/drive.file",
    access_type: "offline",
    prompt: "consent",
    state,
  }).toString();
  return url.href;
}
async function googleToken(env: HostedEnv, params: Record<string, string>) {
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    body: new URLSearchParams({
      client_id: env.GOOGLE_CLIENT,
      client_secret: env.GOOGLE_SECRET,
      ...params,
    }),
    redirect: "manual",
  });
  assert(
    res.ok,
    409,
    "Google authorization expired. Reconnect the original Google account.",
  );
  return z
    .object({ access_token: z.string(), refresh_token: z.string().optional() })
    .parse(await res.json());
}
export async function googleCallback(
  env: HostedEnv,
  customer: string,
  url: URL,
) {
  const data = await consume(
    env,
    "google:" + url.origin,
    url.searchParams.get("state") || "",
  );
  assert(
    data.customer === customer && url.searchParams.get("code"),
    403,
    "Sheet connection was cancelled.",
  );
  const token = await googleToken(env, {
    grant_type: "authorization_code",
    code: url.searchParams.get("code")!,
    redirect_uri: url.origin + "/cloud/google/callback",
  });
  assert(
    token.refresh_token,
    409,
    "Google did not grant offline access. Connect again.",
  );
  const current = await getBook(env);
  if (current?.file_id)
    await sheets(token.access_token, current.file_id + "?fields=spreadsheetId");
  const result = await env.DB.prepare(
    "INSERT INTO hosted_workbook(id,generation,creation_key,refresh_secret) VALUES(1,?,?,?) ON CONFLICT(id) DO UPDATE SET refresh_secret=excluded.refresh_secret,generation=excluded.generation,grant_hash=NULL,error=NULL WHERE hosted_workbook.lease_until<?",
  )
    .bind(id(), id(), await encrypt(env, token.refresh_token), now())
    .run();
  assert(
    result.meta.changes,
    409,
    "A refresh is in progress. Reconnect after it finishes.",
  );
}
async function sheets(token: string, path: string, body?: object) {
  const res = await fetch(
    "https://sheets.googleapis.com/v4/spreadsheets" + (path ? "/" + path : ""),
    {
      method: body ? "POST" : "GET",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
      redirect: "manual",
    },
  );
  assert(
    res.ok,
    res.status === 429 ? 429 : 502,
    res.status === 429
      ? "Google's refresh quota was reached. Retry later."
      : "Google could not update this workbook. Check file access and reconnect.",
  );
  return (await res.json()) as any;
}
export async function workbookRows(ledger: Ledger, year: number) {
  const count = await ledger.db
    .prepare(
      "SELECT COUNT(*) n FROM transactions WHERE deleted=0 AND account_id IN(SELECT id FROM accounts WHERE archived=0)",
    )
    .first<{ n: number }>();
  assert(
    count!.n <= 5000,
    413,
    "The beta workbook supports 5,000 transactions. Your existing workbook is intact.",
  );
  const tx = await ledger.db
    .prepare(
      "SELECT t.id,t.revision,t.posted,t.description,t.amount,t.type,t.merchant,t.notes,t.tags,t.reviewed,t.pending,a.name account,a.currency,c.name category FROM transactions t JOIN accounts a ON a.id=t.account_id LEFT JOIN categories c ON c.id=t.category_id WHERE t.deleted=0 AND a.archived=0 ORDER BY t.posted DESC,t.id",
    )
    .all<any>();
  const accounts = await ledger.accounts();
  const months = await ledger.db
    .prepare(
      "SELECT substr(posted,1,7) month,currency,SUM(CASE WHEN type='income' THEN amount ELSE 0 END) income,SUM(CASE WHEN type='expense' THEN -amount ELSE 0 END) expenses FROM ledger_entries WHERE deleted=0 AND pending=0 AND type IN ('income','expense') AND account_id IN(SELECT id FROM accounts WHERE archived=0) GROUP BY 1,2 ORDER BY 1,2",
    )
    .all<{
      month: string;
      currency: string;
      income: number;
      expenses: number;
    }>();
  const tax = await taxState(ledger, year);
  return [
    tx.results.map((t) => [
      t.id,
      t.revision,
      t.posted,
      t.account,
      t.description,
      t.amount / 100,
      t.currency,
      t.category ?? "",
      t.type,
      t.merchant,
      t.notes,
      t.tags,
      !!t.reviewed,
      !!t.pending,
    ]),
    accounts.map((a) => [
      a.name,
      a.kind,
      a.currency,
      a.balance / 100,
      a.balance_at ?? "",
      a.id,
    ]),
    months.results.map((m) => [
      m.month,
      m.currency,
      m.income / 100,
      m.expenses / 100,
      (m.income - m.expenses) / 100,
    ]),
    (year === 2026 || tax.revision > 0 ? ["federal", "state"] : []).flatMap(
      (j) =>
        (j === "federal"
          ? tax.estimate.federalSchedule
          : tax.estimate.stateSchedule
        ).map((q) => [
          year,
          tax.profile.asOf,
          tax.profile.transactionTag,
          j,
          q.quarter,
          q.due,
          q.target / 100,
          q.shortfall / 100,
        ]),
    ),
    [],
    [
      ["Tax year", year],
      ["Cutoff", tax.profile.asOf],
      ["Transaction tag", tax.profile.transactionTag],
      ["Excluded currencies", tax.excludedCurrencies.join(", ")],
      ...tax.estimate.assumptions.map((a) => ["Calculation", a]),
      ...Object.entries(tax.profile).map(([k, v]) => [
        k,
        typeof v === "object" ? JSON.stringify(v) : String(v),
      ]),
    ],
    ...(await syncedReports(ledger, year)),
  ] as Cell[][][];
}
function row(values: Cell[]) {
  return {
    values: values.map((v) => ({
      userEnteredValue:
        v == null
          ? {}
          : typeof v === "number"
            ? { numberValue: v }
            : typeof v === "boolean"
              ? { boolValue: v }
              : { stringValue: v },
    })),
  };
}
// Drive app properties let a retry recover an already-created file if the original response was lost.
async function findOrCreateFile(token: string, creationKey: string) {
  const endpoint = "https://www.googleapis.com/drive/v3/files";
  const headers = {
    Authorization: "Bearer " + token,
    "Content-Type": "application/json",
  };
  const query = new URLSearchParams({
    q:
      "trashed=false and appProperties has { key='phooeyWorkbook' and value='" +
      creationKey +
      "' }",
    fields: "incompleteSearch,files(id)",
    pageSize: "2",
    spaces: "drive",
  });
  const found = await fetch(endpoint + "?" + query, {
    headers,
    redirect: "manual",
  });
  assert(
    found.ok,
    502,
    "Google Drive could not locate your workbook. Please retry.",
  );
  const result = z
    .object({
      incompleteSearch: z.boolean().optional(),
      files: z.array(z.object({ id: z.string().regex(/^[\w-]+$/) })),
    })
    .parse(await found.json());
  assert(
    !result.incompleteSearch && result.files.length <= 1,
    409,
    "Google Drive could not identify a unique workbook. Contact support.",
  );
  if (result.files.length) return result.files[0].id;
  const created = await fetch(endpoint + "?fields=id", {
    method: "POST",
    headers,
    redirect: "manual",
    body: JSON.stringify({
      name: "phooey! finance · Foundation",
      mimeType: "application/vnd.google-apps.spreadsheet",
      appProperties: { phooeyWorkbook: creationKey },
    }),
  });
  assert(
    created.ok,
    502,
    "Google Drive could not create the workbook. Please retry.",
  );
  return z
    .object({ id: z.string().regex(/^[\w-]+$/) })
    .parse(await created.json()).id;
}
export async function refreshWorkbook(env: HostedEnv) {
  const lease = id();
  const b = await env.DB.prepare(
    "UPDATE hosted_workbook SET lease=?,lease_until=? WHERE id=1 AND lease_until<? RETURNING *",
  )
    .bind(lease, now() + 900, now())
    .first<Book>();
  assert(b, 409, "Connect a workbook or wait for its current refresh.");
  try {
    const data = await workbookRows(
      new Ledger(env.DB, "sheets"),
      b.tax_year ?? new Date().getUTCFullYear(),
    );
    const token = (
      await googleToken(env, {
        grant_type: "refresh_token",
        refresh_token: await decrypt(env, b.refresh_secret),
      })
    ).access_token;
    let file = b.file_id;
    if (!file) {
      file = await findOrCreateFile(token, b.creation_key);
      const saved = await env.DB.prepare(
        "UPDATE hosted_workbook SET file_id=? WHERE id=1 AND lease=?",
      )
        .bind(file, lease)
        .run();
      assert(saved.meta.changes, 409, "The workbook connection changed.");
    }
    const metaPath =
      file +
      "?fields=sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),developerMetadata(metadataKey))";
    const reportMarker = "phooeyFoundationReport";
    let meta = await sheets(token, metaPath);
    const missing = (b.synced_at ? reportTabs : tabs).filter(
      (t) =>
        !meta.sheets.some(
          (s: any) =>
            s.properties.sheetId === t.id || s.properties.title === t.name,
        ),
    );
    if (missing.length) {
      const additions: object[] = [];
      for (const t of missing) {
        additions.push({
          addSheet: {
            properties: {
              sheetId: t.id,
              title: t.name,
              ...(t.id < 120 ? { index: tabs.indexOf(t) } : {}),
              gridProperties: { frozenRowCount: 1 },
            },
          },
        });
        if (t.id >= 120)
          additions.push({
            createDeveloperMetadata: {
              developerMetadata: {
                metadataKey: reportMarker,
                metadataValue: "1",
                visibility: "DOCUMENT",
                location: { sheetId: t.id },
              },
            },
          });
      }
      await sheets(token, file + ":batchUpdate", { requests: additions });
      meta = await sheets(token, metaPath);
    }
    // A same-named report or reserved ID may already belong to the customer.
    // Only write extensions that this publisher created and marked.
    const managedTabs = tabs.filter(
      (t) =>
        t.id < 120 ||
        meta.sheets.some(
          (s: any) =>
            s.properties.sheetId === t.id &&
            s.developerMetadata?.some(
              (m: any) => m.metadataKey === reportMarker,
            ),
        ),
    );
    const headers = await sheets(token, file + "/values:batchGetByDataFilter", {
      dataFilters: managedTabs.map((t) => ({
        gridRange: { sheetId: t.id, startRowIndex: 0, endRowIndex: 1 },
      })),
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    const firstColumn = await sheets(
      token,
      file + "/values:batchGetByDataFilter",
      {
        dataFilters: [
          {
            gridRange: {
              sheetId: 101,
              startColumnIndex: 0,
              endColumnIndex: 1,
              startRowIndex: 1,
            },
          },
        ],
        valueRenderOption: "UNFORMATTED_VALUE",
      },
    );
    const oldIds = (firstColumn.valueRanges?.[0]?.valueRange?.values ??
      []) as Cell[][];
    // Keep each existing transaction on its current row, including after a user sort.
    // A retired ID leaves a blank managed row, so annotations do not move to a different transaction.
    if (oldIds.length) {
      const seen = new Set<string>(),
        byId = new Map(data[0].map((r) => [String(r[0]), r]));
      const arranged: Cell[][] = [];
      for (const [value] of oldIds) {
        const key = String(value ?? "");
        assert(
          !key || !seen.has(key),
          409,
          "A transaction ID appears twice in the feed. Restore its ID column before refreshing.",
        );
        if (key) seen.add(key);
        arranged.push(byId.get(key) ?? (key ? [key] : []));
        byId.delete(key);
      }
      data[0] = arranged.concat([...byId.values()]);
    }
    const requests: object[] = [];
    for (const t of managedTabs) {
      const index = tabs.indexOf(t);
      const props = meta.sheets.find(
        (s: any) => s.properties.sheetId === t.id,
      )?.properties;
      assert(
        props,
        409,
        "A managed worksheet was removed. Restore the tab or disconnect and create a new workbook.",
      );
      const header =
        headers.valueRanges?.find((v: any) =>
          v.dataFilters?.some((f: any) => f.gridRange?.sheetId === t.id),
        )?.valueRange?.values?.[0] ?? [];
      assert(
        ((!b.synced_at || t.id >= 120) && header.length === 0) ||
          t.columns.every((label, i) => header[i] === label),
        409,
        "Managed columns moved or were renamed. Restore their headers and keep custom columns to the right; your workbook has not been overwritten.",
      );
      const rows = data[index] ?? [];
      const end =
        index === 4
          ? 1
          : Math.max(props.gridProperties.rowCount, rows.length + 1);
      if (end > props.gridProperties.rowCount)
        requests.push({
          appendDimension: {
            sheetId: t.id,
            dimension: "ROWS",
            length: end - props.gridProperties.rowCount,
          },
        });
      if (props.gridProperties.columnCount < t.columns.length)
        requests.push({
          appendDimension: {
            sheetId: t.id,
            dimension: "COLUMNS",
            length: t.columns.length - props.gridProperties.columnCount,
          },
        });
      // Bound the update to owned columns. Edits receives only its header, ever.
      requests.push({
        updateCells: {
          range: {
            sheetId: t.id,
            startRowIndex: 0,
            endRowIndex: end,
            startColumnIndex: 0,
            endColumnIndex: t.columns.length,
          },
          rows: [row(t.columns), ...rows.map(row)],
          fields: "userEnteredValue",
        },
      });
      if (!b.synced_at) {
        if (t.id === 101 || t.id === 105) {
          requests.push({
            updateDimensionProperties: {
              range: {
                sheetId: t.id,
                dimension: "COLUMNS",
                startIndex: 0,
                endIndex: 2,
              },
              properties: { hiddenByUser: true },
              fields: "hiddenByUser",
            },
          });
          requests.push({
            updateDimensionProperties: {
              range: {
                sheetId: t.id,
                dimension: "COLUMNS",
                startIndex: 4,
                endIndex: 5,
              },
              properties: { pixelSize: 240 },
              fields: "pixelSize",
            },
          });
        }
        if (t.id === 101)
          requests.push({
            setBasicFilter: {
              filter: {
                range: {
                  sheetId: t.id,
                  startRowIndex: 0,
                  startColumnIndex: 0,
                  endColumnIndex: t.columns.length,
                },
              },
            },
          });
        for (const [i, label] of t.columns.entries()) {
          if (
            [
              "Amount",
              "Balance",
              "Income",
              "Expenses",
              "Net",
              "Estimate (USD)",
              "Shortfall (USD)",
            ].includes(label)
          ) {
            requests.push({
              repeatCell: {
                range: {
                  sheetId: t.id,
                  startRowIndex: 1,
                  startColumnIndex: i,
                  endColumnIndex: i + 1,
                },
                cell: {
                  userEnteredFormat: {
                    numberFormat: {
                      type: "NUMBER",
                      pattern: "#,##0.00;(#,##0.00);0.00",
                    },
                  },
                },
                fields: "userEnteredFormat.numberFormat",
              },
            });
          }
        }
      }
      requests.push({
        repeatCell: {
          range: {
            sheetId: t.id,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: t.columns.length,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.16, green: 0.39, blue: 0.29 },
              textFormat: {
                bold: true,
                foregroundColor: { red: 1, green: 1, blue: 1 },
              },
            },
          },
          fields: "userEnteredFormat",
        },
      });
    }
    await sheets(token, file + ":batchUpdate", { requests });
    await env.DB.prepare(
      "UPDATE hosted_workbook SET synced_at=?,error=NULL WHERE lease=?",
    )
      .bind(now(), lease)
      .run();
  } catch (error) {
    const message = errorMessage(error);
    await env.DB.prepare("UPDATE hosted_workbook SET error=? WHERE lease=?")
      .bind(message.slice(0, 240), lease)
      .run();
    throw error;
  } finally {
    await env.DB.prepare(
      "UPDATE hosted_workbook SET lease=NULL,lease_until=0 WHERE lease=?",
    )
      .bind(lease)
      .run();
  }
}
export async function disconnectBook(env: HostedEnv) {
  const b = await env.DB.prepare(
    "DELETE FROM hosted_workbook WHERE id=1 AND lease_until<? RETURNING refresh_secret",
  )
    .bind(now())
    .first<{ refresh_secret: string }>();
  assert(b, 409, "Wait for the current refresh before disconnecting.");
  const res = await fetch("https://oauth2.googleapis.com/revoke", {
    method: "POST",
    body: new URLSearchParams({ token: await decrypt(env, b.refresh_secret) }),
  });
  assert(
    res.ok || res.status === 400,
    502,
    "Phooey disconnected. Remove the remaining Google permission in your Google Account.",
  );
}
export async function pair(env: HostedEnv, customer: string) {
  const b = await getBook(env);
  assert(b?.file_id, 409, "Create the workbook first.");
  return challenge(
    env,
    "pair:" + env.WORKSPACE_ID,
    { customer, file: b.file_id, generation: b.generation },
    300,
  );
}
export async function finishPair(
  env: HostedEnv,
  code: string,
  file: string,
  customer: string,
) {
  const grant = await consume(env, "pair:" + env.WORKSPACE_ID, code);
  assert(
    grant.file === file && grant.customer === customer,
    403,
    "Use the original connected workbook.",
  );
  const token = randomToken();
  const r = await env.DB.prepare(
    "UPDATE hosted_workbook SET grant_hash=? WHERE file_id=? AND generation=?",
  )
    .bind(await hash(token), file, grant.generation)
    .run();
  assert(r.meta.changes, 409, "Workbook connection changed.");
  return token;
}
export async function checkGrant(env: HostedEnv, token: string, file: string) {
  const b = await getBook(env);
  assert(
    token && b?.grant_hash === (await hash(token)) && b.file_id === file,
    401,
    "Pair this workbook again.",
  );
}
