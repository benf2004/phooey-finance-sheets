# phooey! finance for Google Sheets

[Make a copy of the Foundation workbook](https://docs.google.com/spreadsheets/d/1jGqbCx49IiIUT_LKpVL6osxdALSDF0SEssz4osHbJM8/copy). The public example contains fictional records. You can use the template on its own or bring balances and transactions into your own copy from [phooey! finance](https://github.com/benf2004/phooey-finance).

The workbook includes monthly and yearly budgets, Spending Trends, Estimated Quarterly Tax, a Business Dashboard and Quarterly Earnings. Review the filing year, tax assumptions and Business category tags in your copy.

## Connect a ledger

1. Start with a working copy of your workbook. Choose a first sync date after the last imported transaction and balance in the accounts you intend to connect.
2. In Finance, open Settings → Connections → Create access token. Select only `finance:read`.
3. Create a private standalone Apps Script project. Add the two script files in `connector/` and replace the manifest with its `appsscript.json`. To see the manifest, enable it in Project Settings. A custom Google Cloud project also needs the Sheets API enabled.
4. Set these script properties in Project Settings. Editors of this private script can read them.

| Property | Value |
| --- | --- |
| WORKBOOK_URL | Full Google Sheets URL of your working copy |
| PHOOEY_ORIGIN | HTTPS address of your ledger, without a trailing slash |
| READ_TOKEN | Your read-only Finance access token |
| FIRST_SYNC_DATE | First import date, in YYYY-MM-DD format |

5. Run `prepareWorkbook`. On the Phooey Account Map tab, enter an existing Balance History account ID for each USD account, enter NEW for a new account, or leave it blank to skip it.
6. Run `previewSync` and check the counts in the execution log. Run `syncWorkbook` to apply. Optional daily imports are enabled with `enableDailySync` and stopped with `disableDailySync`.

For an existing compatible workbook, run `addBusinessDashboard` or `addQuarterlyEarnings` to install a missing report. The connector writes feed data into Transactions, Categories and Balance History. Keep your spreadsheet copy and script properties private.

## Development

Use Node 24 or later. Run `npm ci` and `npm test`. The archive preserves source from the retired hosted Sheets experiment for reference. Its sidebar depended on a Finance-hosted backend. Use the standalone connector above for new installations. There is no deployment step for this repository.
