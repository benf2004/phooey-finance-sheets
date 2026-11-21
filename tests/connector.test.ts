import { runInNewContext } from "node:vm";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const script = readFileSync("connector/Code.gs", "utf8");
const ctx: any = {
  Date,
  Utilities: {
    formatDate(d: Date, _zone: string, pattern: string) {
      return pattern === "HH:mm:ss"
        ? d.toISOString().slice(11, 19)
        : d.toISOString().slice(0, 10);
    },
  },
};
runInNewContext(script, ctx);
function grid(id: number, header: string[], rows: any[][] = []) {
  return {
    id,
    header,
    values: [header, ...rows],
    capacity: 100,
    next: rows.length + 1,
    formulas: [] as any[][],
    tz: "UTC",
  };
}
function fixture() {
  const tx = grid(1, [
    "Date",
    "Description",
    "Amount",
    "Category",
    "Private note",
    "Account ID",
    "Transaction ID",
    "Month",
    "Week",
  ]);
  const cats = grid(
    2,
    ["Category", "Group", "Type", "Tags"],
    [["Office", "Work", "Expense", "Business"]],
  );
  const hist = grid(
    3,
    ["Date", "Account", "Account ID", "Balance ID", "Balance", "Class", "Time"],
    [[45650, "Corporate card", "tiller-cc", "original", 30, "Liability", ""]],
  );
  const rules = grid(
    4,
    ["Category", "Description Contains", "Amount Max"],
    [["Office", "printer", 0]],
  );
  const mapping = grid(
    5,
    ["Phooey ID", "Phooey account", "Currency", "Tiller Account ID or NEW"],
    [["card", "Card", "USD", "tiller-cc"]],
  );
  const accounts = [
    {
      id: "card",
      name: "Corporate card",
      kind: "credit",
      currency: "USD",
      balance: -3000,
      balance_at: "2026-06-01T16:00:00Z",
    },
  ];
  const t = {
    id: "sample",
    part: 0,
    amount: -1000,
    tags: ["Business"],
    type: "expense",
    account_id: "card",
    category_id: null,
    description: "=Printer",
    posted: "2026-05-31",
  };
  const balances = [
    { account_id: "card", balance: -7000, captured_at: "2026-06-01T11:00:00Z" },
  ];
  return {
    tables: { tx, cats, hist, rules },
    mapping,
    data: {
      accounts,
      categories: [],
      balances,
      transactions: [t, { ...t, part: 1, amount: -1500, description: "Other" }],
    },
  };
}
function apply(tables: ReturnType<typeof fixture>["tables"], requests: any[]) {
  for (const { updateCells: r } of requests) {
    if (!r) continue;
    const grid = Object.values(tables).find((t) => t.id === r.start.sheetId)!;
    const row =
      grid.values[r.start.rowIndex] ?? (grid.values[r.start.rowIndex] = []);
    r.rows[0].values.forEach(
      ({ userEnteredValue: v }: any, i: number) =>
        (row[r.start.columnIndex + i] = v?.stringValue ?? v?.numberValue ?? ""),
    );
  }
}
describe("native Tiller feeds", () => {
  it("keeps sheet metadata, maps split amounts, applies AutoCat, and uses the latest daily liability balance", () => {
    const { tables, mapping, data } = fixture();
    const run = () => ctx.plan_(tables, mapping, data, "2026-01-01", "UTC");
    apply(tables, run().requests);
    expect(tables.tx.values[1]).toMatchObject({
      1: "=Printer",
      2: -10,
      3: "Office",
      5: "tiller-cc",
      6: "phooey:tx:sample:0",
      7: 46143,
      8: 46173,
    });
    expect(tables.tx.values[2][2]).toBe(-15);
    expect(tables.cats.values.slice(1)).toEqual([
      ["Office", "Work", "Expense", "Business"],
      ["Phooey business expense", "Business", "Expense", "Business"],
    ]);
    expect(tables.hist.values[2].slice(1, 6)).toEqual([
      "Corporate card",
      "tiller-cc",
      "phooey:balance:card:2026-06-01",
      30,
      "Liability",
    ]);
    tables.tx.values[1][4] = "Customer note";
    tables.tx.values[1][3] = "Locally reviewed";
    tables.cats.values.push([
      "Locally reviewed",
      "Work",
      "Expense",
      "Business",
    ]);
    tables.cats.next++;
    tables.tx.values = [
      tables.tx.header,
      ...tables.tx.values.slice(1).reverse(),
    ];
    expect(run().requests).toEqual([]);
    data.transactions.pop();
    apply(tables, run().requests);
    expect(
      tables.tx.values.find((r) => r[6] === "phooey:tx:sample:0")?.[4],
    ).toBe("Customer note");
    expect(tables.tx.values.some((r) => r[6] === "phooey:tx:sample:1")).toBe(
      false,
    );
  });
  it("rejects overlapping Tiller feeds and ambiguous account mappings", () => {
    const { tables, mapping, data } = fixture();
    tables.hist.values[1][0] = ctx.serial_("2026-06-01");
    expect(() => ctx.plan_(tables, mapping, data, "2026-01-01", "UTC")).toThrow(
      /overlap/,
    );
    mapping.values[1][3] = "unknown";
    expect(() => ctx.plan_(tables, mapping, data, "2026-01-01", "UTC")).toThrow(
      /missing/,
    );
  });
});
