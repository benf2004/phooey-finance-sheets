import { expect, it } from "vitest";
import { runInNewContext } from "node:vm";
import { readFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
it("submits pasted drafts from the original workbook and keeps the request key after an interrupted save", async () => {
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
    "Request key",
  ];
  const rows: any[][] = [
    columns,
    ...["First", "Second"].map((d) => [
      "",
      "",
      "2026-11-03",
      "Checking",
      d,
      -12.5,
      "USD",
      "Office",
      "expense",
      "",
      "",
      '["Business"]',
      true,
      false,
      "",
    ]),
  ];
  const active = { getRow: () => 2, getNumRows: () => 2 };
  const sheet = {
    getSheetId: () => 105,
    getActiveRange: () => active,
    getRange: (r: number, c: number, height = 1, width = 1) => ({
      getValues: () =>
        rows
          .slice(r - 1, r - 1 + height)
          .map((line) => line.slice(c - 1, c - 1 + width)),
      setValue: (v: unknown) => {
        rows[r - 1][c - 1] = v;
      },
      setValues: (v: any[][]) => {
        rows[r - 1].splice(c - 1, width, ...v[0]);
      },
    }),
  };
  let locked = false;
  const context: any = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getActiveSheet: () => sheet,
        getSheets: () => [sheet],
      }),
      flush() {},
    },
    LockService: {
      getDocumentLock: () => ({
        waitLock() {
          locked = true;
        },
        releaseLock() {
          locked = false;
        },
      }),
    },
    Utilities: { getUuid: randomUUID },
  };
  runInNewContext(await readFile("archive/addon/Code.gs", "utf8"), context);
  const saved: any[] = [];
  context.api_ = (path: string, data: any) => {
    if (path === "options")
      return {
        accounts: [{ id: "a", name: "Checking", currency: "USD" }],
        categories: [{ id: "office", name: "Office" }],
      };
    if (data.transaction.description === "Second")
      throw Error("Connection interrupted");
    saved.push(data);
    return { id: "transaction-a", revision: 1 };
  };
  expect(() => context.saveDraft()).toThrow(
    "Stopped at row 3. 1 earlier rows saved.",
  );
  expect(locked).toBe(false);
  expect(saved[0].transaction).toMatchObject({
    account_id: "a",
    amount: -1250,
    category_id: "office",
    tags: ["Business"],
  });
  expect(rows[1].slice(0, 2)).toEqual(["transaction-a", 1]);
  const pendingKey = rows[2][14];
  expect(pendingKey).toBeTruthy();
  active.getRow = () => 3;
  active.getNumRows = () => 1;
  context.api_ = (path: string, data: any) => {
    if (path === "options")
      return {
        accounts: [{ id: "a", name: "Checking", currency: "USD" }],
        categories: [{ id: "office", name: "Office" }],
      };
    expect(data.requestKey).toBe(pendingKey);
    return { id: "transaction-b", revision: 1 };
  };
  expect(context.saveDraft()).toMatch(/1 draft/);
  expect(rows[2][0]).toBe("transaction-b");
  columns[4] = "Personal notes";
  expect(() => context.saveDraft()).toThrow("original column order");
});
