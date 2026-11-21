// Original phooey! finance Google Sheets templates. Run createFoundationWorkbook
// in a private standalone Apps Script project with the advanced Sheets service.
const foundationNames = [
  "Spending Trends",
  "Transactions",
  "Monthly Budget",
  "Yearly Budget",
  "Categories",
  "Budget Plan",
  "Accounts",
  "Balances",
  "Balance History",
  "AutoCat",
  "Transaction Tracker",
  "Business Dashboard",
  "Quarterly Earnings",
  "Est Quarterly Tax",
  "Archived Est Quarterly Tax",
  "Foundation Activity",
];
const foundationHeaders = {
  Transactions: [
    "Date",
    "Description",
    "Category",
    "Amount",
    "Account",
    "Account ID",
    "Transaction ID",
    "Account #",
    "Institution",
    "Month",
    "Week",
    "Full Description",
    "Date Added",
  ],
  Categories: ["Category", "Group", "Type", "Tags", "Hide From Reports"],
  "Balance History": [
    "Date",
    "Time",
    "Account",
    "Account #",
    "Account ID",
    "Balance ID",
    "Institution",
    "Balance",
    "Month",
    "Week",
    "Type",
    "Class",
    "Account Status",
    "Date Added",
  ],
  AutoCat: [
    "Category",
    "Description Contains",
    "Account Contains",
    "Institution Contains",
    "Amount Min",
    "Amount Max",
  ],
  "Budget Plan": ["Year", "Month (1–12)", "Category", "Target (USD)"],
  Accounts: ["Account ID", "Display name override", "Group"],
  "Foundation Activity": [
    "Date",
    "Description",
    "Category",
    "Amount",
    "Account",
    "Type",
    "Business",
    "Account ID",
  ],
};
function foundationTemplates(year) {
  const out = {};
  function sheet(name) {
    const s = {
      name,
      rows: [],
      widths: [170, 170, 160, 160, 160, 155, 170, 160],
      money: [],
      dates: [],
    };
    out[name] = s;
    return s;
  }
  function place(s, r, c, ...values) {
    const row = s.rows[r - 1] || (s.rows[r - 1] = []);
    values.forEach((v, i) => (row[c + i - 1] = v));
  }
  for (const name of foundationNames) {
    const s = sheet(name);
    if (foundationHeaders[name]) place(s, 1, 1, ...foundationHeaders[name]);
    else place(s, 1, 1, name);
  }
  // Find feed columns by their headers so a Tiller copy can have a different order.
  const column = (tab, key) => {
    const letter = `SUBSTITUTE(ADDRESS(1,MATCH("${key}",'${tab}'!1:1,0),4),"1","")`;
    return `INDIRECT("'${tab}'!"&${letter}&"2:"&${letter})`;
  };

  const tx = (key) => column("Transactions", key),
    cat = (key) => column("Categories", key);
  const a = out["Foundation Activity"];
  place(
    a,
    2,
    1,
    "=IFNA(FILTER({" +
      ["Date", "Description", "Category", "Amount", "Account"]
        .map(tx)
        .join(",") +
      ",ARRAYFORMULA(IFNA(VLOOKUP(" +
      tx("Category") +
      ", {" +
      cat("Category") +
      "," +
      cat("Type") +
      '},2,FALSE),"Uncategorized")),ARRAYFORMULA(REGEXMATCH(LOWER(IFNA(VLOOKUP(' +
      tx("Category") +
      ", {" +
      cat("Category") +
      "," +
      cat("Tags") +
      '},2,FALSE),"")),"(^|[,;[:space:]])business($|[,;[:space:]])")),' +
      tx("Account ID") +
      "},ISNUMBER(" +
      tx("Date") +
      ')),"")',
  );
  a.dates = ["A2:A"];
  a.money = ["D2:D"];
  out.Transactions.dates = ["A2:A", "J2:K", "M2:M"];
  out.Transactions.money = ["D2:D"];
  out["Balance History"].dates = ["A2:A", "I2:J", "N2:N"];
  out["Balance History"].money = ["H2:H"];
  out["Budget Plan"].money = ["D2:D"];
  const src = "'Foundation Activity'!",
    col = (l) => src + l + "2:" + l;
  const yearsel = (s) => place(s, 2, 1, "Report year", year);
  const actual = (type, start, end, biz) =>
    (type === "Expense" ? "-" : "") +
    "SUMIFS(" +
    col("D") +
    "," +
    col("F") +
    ',"' +
    type +
    '",' +
    col("A") +
    ',">="&' +
    start +
    "," +
    col("A") +
    ',"<"&' +
    end +
    (biz ? "," + col("G") + ",TRUE" : "") +
    ")";
  for (const [name, biz, quarter] of [
    ["Spending Trends", false, false],
    ["Business Dashboard", true, false],
    ["Quarterly Earnings", true, true],
  ]) {
    const s = out[name];
    yearsel(s);
    place(
      s,
      4,
      1,
      quarter ? "Calendar quarter" : "Month",
      "Income",
      "Expenses",
      biz ? "Net profit" : "Net cash flow",
    );
    for (let i = 0; i < (quarter ? 4 : 12); i++) {
      const r = 5 + i,
        m = 1 + i * (quarter ? 3 : 1),
        start = "DATE($B$2," + m + ",1)",
        end = "EDATE(" + start + "," + (quarter ? 3 : 1) + ")";
      place(
        s,
        r,
        1,
        quarter ? "Q" + (i + 1) : "=" + start,
        "=" + actual("Income", start, end, biz),
        "=" + actual("Expense", start, end, biz),
        "=B" + r + "-C" + r,
      );
    }
    const last = quarter ? 8 : 16;
    place(
      s,
      last + 1,
      1,
      "Year total",
      ...["B", "C", "D"].map((c) => "=SUM(" + c + "5:" + c + last + ")"),
    );
    s.money = ["B5:D" + (last + 1)];
    if (!quarter) s.months = ["A5:A16"];
    if (name === "Spending Trends") {
      place(s, 20, 1, "Current net worth", "=SUM(Balances!E5:E)");
      s.money.push("B20");
    }
  }
  for (const monthly of [true, false]) {
    const s = out[monthly ? "Monthly Budget" : "Yearly Budget"];
    yearsel(s);
    if (monthly) place(s, 3, 1, "Month (1–12)", 1);
    place(
      s,
      4,
      1,
      "Category",
      "Type",
      "Budget",
      "Actual",
      "Favorable variance",
    );
    place(
      s,
      5,
      1,
      '=IFNA(FILTER({Categories!A2:A,Categories!C2:C},Categories!A2:A<>"",Categories!C2:C<>"Transfer"),"")',
    );
    const start = monthly ? "DATE($B$2,$B$3,1)" : "DATE($B$2,1,1)",
      end = monthly ? "EDATE(" + start + ",1)" : "DATE($B$2+1,1,1)";
    place(
      s,
      5,
      3,
      "=IFNA(MAP(FILTER(A5:A,A5:A<>\"\"),LAMBDA(k,SUMIFS('Budget Plan'!D2:D,'Budget Plan'!A2:A,$B$2,'Budget Plan'!C2:C,k" +
        (monthly ? ",'Budget Plan'!B2:B,$B$3" : "") +
        '))),"")',
      '=IFNA(MAP(FILTER(A5:A,A5:A<>""),FILTER(B5:B,A5:A<>""),LAMBDA(k,t,IF(t="Expense",-1,1)*SUMIFS(' +
        col("D") +
        "," +
        col("C") +
        ",k," +
        col("A") +
        ',">="&' +
        start +
        "," +
        col("A") +
        ',"<"&' +
        end +
        '))),"")',
      '=ARRAYFORMULA(IF(A5:A="","",IF(B5:B="Expense",C5:C-D5:D,D5:D-C5:C)))',
    );
    s.money = ["C5:E"];
  }
  const tracker = out["Transaction Tracker"];
  place(tracker, 2, 1, "Account filter", "", "Category filter", "");
  place(tracker, 3, 1, "Start date", "", "End date", "");
  place(tracker, 4, 1, "Date", "Description", "Category", "Amount", "Account");
  place(
    tracker,
    5,
    1,
    "=IFNA(FILTER(" +
      src +
      "A2:E,ISNUMBER(" +
      col("A") +
      '),IF(B2="",ISNUMBER(' +
      col("A") +
      ")," +
      col("E") +
      '=B2),IF(D2="",ISNUMBER(' +
      col("A") +
      ")," +
      col("C") +
      "=D2)," +
      col("A") +
      '>=IF(B3="",0,B3),' +
      col("A") +
      '<=IF(D3="",DATE(9999,12,31),D3)),"")',
  );
  tracker.dates = ["A5:A", "B3", "D3"];
  tracker.money = ["D5:D"];
  const balances = out.Balances;
  place(balances, 2, 1, "Current net worth", "=SUM(E5:E)");
  place(
    balances,
    4,
    1,
    "Account ID",
    "Account",
    "Class",
    "Group",
    "Net balance",
    "Last update",
  );
  const h = "'Balance History'!",
    latest =
      "SORT(FILTER({" +
      h +
      "E2:E," +
      h +
      "C2:C," +
      h +
      "L2:L," +
      h +
      "H2:H," +
      h +
      "A2:A," +
      h +
      "B2:B}," +
      h +
      'E2:E<>""),5,FALSE,6,FALSE)';
  place(
    balances,
    5,
    1,
    "=IFNA(LET(obs," +
      latest +
      ',ids,UNIQUE(INDEX(obs,0,1)),labels,ARRAYFORMULA(IFNA(VLOOKUP(ids,Accounts!A2:C,2,FALSE),"")),classes,ARRAYFORMULA(VLOOKUP(ids,obs,3,FALSE)),{ids,ARRAYFORMULA(IF(labels="",VLOOKUP(ids,obs,2,FALSE),labels)),classes,ARRAYFORMULA(IFNA(VLOOKUP(ids,Accounts!A2:C,3,FALSE),"")),ARRAYFORMULA(VLOOKUP(ids,obs,4,FALSE)*IF(classes="Liability",-1,1)),ARRAYFORMULA(VLOOKUP(ids,obs,5,FALSE))}),"")',
  );
  balances.money = ["B2", "E5:E"];
  balances.dates = ["F5:F"];
  for (const previous of [false, true]) {
    const s =
        out[previous ? "Archived Est Quarterly Tax" : "Est Quarterly Tax"],
      y = previous ? year - 1 : year;
    place(s, 2, 1, "Tax year", y);
    place(
      s,
      3,
      1,
      "Rate schedule year",
      y,
      "Review inputs when changing tax year.",
    );
    place(
      s,
      4,
      1,
      "US tax period",
      "Income",
      "Expenses",
      "Net profit",
      "Cumulative profit",
      "SE tax to date",
      "Taxable income",
      "Federal total to date",
      "Federal for period",
      "State for period",
    );
    const cuts = [1, 4, 6, 9, 13];
    for (let i = 0; i < 4; i++) {
      const r = 5 + i,
        start = "DATE($B$2," + cuts[i] + ",1)",
        end = "DATE($B$2," + cuts[i + 1] + ",1)";
      const brackets = Array.from(
        { length: 7 },
        (_, b) =>
          "MAX(0,G" +
          r +
          "-A" +
          (17 + b) +
          ")*" +
          (b ? "(B" + (17 + b) + "-B" + (16 + b) + ")" : "B17"),
      ).join("+");
      place(
        s,
        r,
        1,
        ["Jan–Mar", "Apr–May", "Jun–Aug", "Sep–Dec"][i],
        "=" + actual("Income", start, end, true),
        "=" + actual("Expense", start, end, true),
        "=B" + r + "-C" + r,
        "=SUM(D$5:D" + r + ")",
        "=ROUND(MIN(MAX(0,E" +
          r +
          ")*0.9235,$B$12)*0.124+MAX(0,E" +
          r +
          ")*0.9235*0.029,0)",
        "=MAX(0,E" + r + "-ROUND(F" + r + "/2,0)-$B$11*" + (i + 1) + "/4)",
        '=IF($B$2<>$B$3,"",ROUND(' +
          brackets +
          ",0)+F" +
          r +
          "+ROUND(MAX(0,MAX(0,E" +
          r +
          ")*0.9235-$B$13)*0.009,0))",
        "=IF(H" + r + '="","",H' + r + (i ? "-H" + (r - 1) : "") + ")",
        '=IF($B$2<>$B$3,"",(MAX(0,E' +
          r +
          ")" +
          (i ? "-MAX(0,E" + (r - 1) + ")" : "") +
          ")*$B$14)",
      );
    }
    place(s, 10, 1, "Single filer assumptions");
    place(s, 11, 1, "Standard deduction", previous ? 15750 : 16100);
    place(s, 12, 1, "Social Security wage base", previous ? 176100 : 184500);
    place(s, 13, 1, "Additional Medicare threshold", 200000);
    place(s, 14, 1, "Flat state rate", 0);
    place(s, 16, 1, "Federal band starts", "Marginal rate");
    (previous
      ? [0, 11925, 48475, 103350, 197300, 250525, 626350]
      : [0, 12400, 50400, 105700, 201775, 256225, 640600]
    ).forEach((n, i) =>
      place(s, 17 + i, 1, n, [0.1, 0.12, 0.22, 0.24, 0.32, 0.35, 0.37][i]),
    );
    place(
      s,
      25,
      1,
      "USD business category income and expenses. Single filer, no other wages, income, credits or withholding. Use the app tax planner for a full profile and payment tracking.",
    );
    s.money = ["B5:J8", "B11:B13", "A17:A23"];
    s.percent = ["B14", "B17:B23"];
    s.widths = [245, 150, 150, 150, 165, 160, 160, 180, 160, 160];
  }
  return out;
}
function writeFoundationTab_(tab, s) {
  const name = s.name,
    width = Math.max(...Array.from(s.rows, (r) => r?.length || 0));
  const values = Array.from({ length: s.rows.length }, (_, i) =>
    Array.from({ length: width }, (_, j) => s.rows[i]?.[j] ?? ""),
  );
  tab.getRange(1, 1, values.length, width).setValues(values);
  tab.setFrozenRows(1);
  tab
    .getRange(1, 1, 1, width)
    .setBackground("#173c35")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  tab.setRowHeight(1, 36);
  if (!foundationHeaders[name])
    tab.getRange(4, 1, 1, width).setBackground("#dae8da").setFontWeight("bold");
  for (const [i, w] of s.widths.entries()) tab.setColumnWidth(i + 1, w);
  for (const rg of s.money)
    tab.getRange(rg).setNumberFormat("#,##0.00;(#,##0.00);0.00");
  for (const rg of s.dates) tab.getRange(rg).setNumberFormat("yyyy-mm-dd");
  for (const rg of s.months || []) tab.getRange(rg).setNumberFormat("mmmm");
  for (const rg of s.percent || []) tab.getRange(rg).setNumberFormat("0.0%");
  if (
    ["Spending Trends", "Business Dashboard", "Quarterly Earnings"].includes(
      name,
    )
  ) {
    const end = name === "Quarterly Earnings" ? 8 : 16;
    tab.getRange("F1").setValue("phooey! finance").setFontWeight("bold");
    tab.insertChart(
      tab
        .newChart()
        .asColumnChart()
        .addRange(tab.getRange("A4:C" + end))
        .setNumHeaders(1)
        .setPosition(4, 6, 0, 0)
        .setOption("legend", { position: "bottom" })
        .setOption("colors", ["#347762", "#d78154"])
        .build(),
    );
  }
}
function writeFoundationTemplates_(book, year) {
  const templates = foundationTemplates(year);
  for (const name of foundationNames)
    if (!book.getSheetByName(name)) book.insertSheet(name);
  for (const name of foundationNames)
    writeFoundationTab_(book.getSheetByName(name), templates[name]);
  book.getSheetByName("Foundation Activity").hideSheet();
}
// Install only missing report tabs in a workbook selected by its owner.
function addFoundationReport_(name) {
  const url =
    PropertiesService.getScriptProperties().getProperty("WORKBOOK_URL");
  if (!url)
    throw Error(
      "Set WORKBOOK_URL to the Google Sheets copy you want to extend.",
    );
  const book = SpreadsheetApp.openByUrl(url),
    templates = foundationTemplates(2026);
  if (book.getSheetByName(name)) return;
  for (const [tab, labels] of [
    [
      "Transactions",
      ["Date", "Description", "Category", "Amount", "Account", "Account ID"],
    ],
    ["Categories", ["Category", "Type", "Tags"]],
  ]) {
    const existing = book.getSheetByName(tab);
    const headers =
      existing
        ?.getRange(1, 1, 1, existing.getLastColumn())
        .getDisplayValues()[0] || [];
    if (!labels.every((label) => headers.includes(label)))
      throw Error(
        "Check the required headers on " +
          tab +
          " before installing this report.",
      );
  }
  let activity = book.getSheetByName("Foundation Activity");
  if (activity && !activity.getRange("A2").getFormula().includes("REGEXMATCH"))
    throw Error("The Foundation Activity name is in use by another tab.");
  if (!activity) {
    activity = book.insertSheet("Foundation Activity");
    writeFoundationTab_(activity, templates["Foundation Activity"]);
    activity.hideSheet();
  }
  writeFoundationTab_(book.insertSheet(name), templates[name]);
}
function addBusinessDashboard() {
  addFoundationReport_("Business Dashboard");
}
function addQuarterlyEarnings() {
  addFoundationReport_("Quarterly Earnings");
}
function createFoundationWorkbook() {
  const book = SpreadsheetApp.create("phooey! finance · Foundation");
  book.getSheets()[0].setName("Spending Trends");
  writeFoundationTemplates_(book, 2026);
  console.log(book.getUrl());
  return book.getUrl();
}
