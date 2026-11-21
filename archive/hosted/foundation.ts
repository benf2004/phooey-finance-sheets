import {
  foundationReport,
  monthNames,
  type ViewId,
} from "../shared/foundation";
import { taxState } from "../server/taxes";
import { foundationData } from "../server/foundation";
import type { Ledger } from "../server/ledger";

const definitions: [ViewId, string, string[]][] = [
  [
    "spending-trends",
    "Spending Trends",
    ["Currency", "Month", "Income", "Expenses", "Net cash flow"],
  ],
  [
    "monthly-budget",
    "Monthly Budget",
    [
      "Currency",
      "Month",
      "Category",
      "Type",
      "Budget",
      "Actual",
      "Favorable variance",
    ],
  ],
  [
    "yearly-budget",
    "Yearly Budget",
    ["Currency", "Category", "Type", "Budget", "Actual", "Favorable variance"],
  ],
  ["categories", "Categories", ["Category", "Type", "Tax treatment"]],
  [
    "accounts",
    "Accounts",
    ["Currency", "Account", "Institution", "Account type"],
  ],
  ["balances", "Balances", ["Currency", "Account", "Class", "Net balance"]],
  [
    "balance-history",
    "Balance History",
    ["Currency", "Observed", "Account", "Balance"],
  ],
  [
    "autocat",
    "AutoCat",
    ["Priority", "Rule", "Match field", "Contains", "Category", "Status"],
  ],
  [
    "transaction-tracker",
    "Transaction Tracker",
    ["Currency", "Date", "Description", "Category", "Account", "Amount"],
  ],
  [
    "business-dashboard",
    "Business Dashboard",
    ["Currency", "Month", "Income", "Expenses", "Net profit"],
  ],
  [
    "quarterly-earnings",
    "Quarterly Earnings",
    ["Currency", "Calendar quarter", "Income", "Expenses", "Net profit"],
  ],
  [
    "archived-tax",
    "Archived Est Quarterly Tax",
    [
      "Year",
      "Quarter",
      "Federal due",
      "Federal estimate (USD)",
      "State estimate (USD)",
    ],
  ],
];
export const reportTabs = definitions.map(([, name, columns], i) => ({
  id: 120 + i,
  name,
  columns,
}));
export async function syncedReports(ledger: Ledger, year: number) {
  const d = await foundationData(ledger, year),
    currencies = [...new Set(d.accounts.map((a) => a.currency))].sort();
  const prior = await taxState(ledger, year - 1);
  return definitions.map(([view]) => {
    if (view === "archived-tax")
      return !prior.revision && year - 1 !== 2026
        ? []
        : prior.estimate.federalSchedule.map((q, i) => [
            year - 1,
            q.quarter,
            q.due,
            q.target / 100,
            prior.estimate.stateSchedule[i].target / 100,
          ]);
    if (view === "categories" || view === "autocat")
      return foundationReport(d, view, "USD", 0).rows;
    return currencies.flatMap((cur) => {
      const rowsFor = (month: number) => {
        const r = foundationReport(d, view, cur, month);
        return r.rows.map((cells) => {
          const values = cells.map((v, i) =>
            r.money.includes(i) ? Number(v) / 100 : v,
          );
          if (view === "accounts") values.pop();
          return [
            cur,
            ...(view === "monthly-budget" ? [monthNames[month - 1]] : []),
            ...values,
          ];
        });
      };
      return view === "monthly-budget"
        ? monthNames.flatMap((_, i) => rowsFor(i + 1))
        : rowsFor(0);
    });
  });
}
