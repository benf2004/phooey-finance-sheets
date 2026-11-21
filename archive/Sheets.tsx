import { useState } from "react";
import { useWorkspace } from "../context";
import { api, post } from "../api";
import { useAsync, Panel, Button, ErrorNotice } from "../ui";
import { BankConnection, type BankConnectionState } from "./BankConnection";
import ImportDialog from "./ImportDialog";
import Categories from "./Categories";
import Taxes from "./Taxes";
import TillerImport from "./TillerImport";
type Status = {
  year: number;
  connected: boolean;
  url: string | null;
  synced: number | null;
  error: string | null;
  paired: boolean;
  addon: string | null;
};
async function cloud<T>(
  action: string,
  write = false,
  body?: object,
): Promise<T> {
  const r = await fetch("/cloud/sheets/" + action, {
      method: write ? "POST" : "GET",
      body: body ? JSON.stringify(body) : undefined,
    }),
    data = (await r.json()) as T & { error?: string };
  if (!r.ok) throw Error(data.error || "Request failed.");
  return data;
}
export default function Sheets() {
  const { version, refresh, accounts } = useWorkspace(),
    state = useAsync(() => cloud<Status>("status"), [version]),
    bank = useAsync(
      () => api<{ connection: BankConnectionState | null }>("/sync"),
      [version],
    );
  const [error, setError] = useState(""),
    [busy, setBusy] = useState(false),
    [code, setCode] = useState(""),
    [csv, setCsv] = useState(false),
    [account, setAccount] = useState(""),
    [panel, setPanel] = useState(""),
    [year, setYear] = useState("");
  const run = async (fn: () => Promise<void>) => {
    setBusy(true);
    setError("");
    try {
      await fn();
      refresh();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };
  const [matches, setMatches] = useState<
    | {
        imported: string;
        bank: string;
        imported_description: string;
        bank_description: string;
        posted: string;
        amount: number;
        account: string;
        currency: string;
      }[]
    | null
  >(null);
  const b = state.data;
  return (
    <main className="content" style={{ maxWidth: 980, margin: "auto" }}>
      <header className="actions">
        <img
          src="/business/assets/wordmark.svg"
          alt="phooey! finance"
          width="145"
        />
        <a href="/cloud/account">Subscription</a>
        <a href="/cloud/enter?view=dashboard">Open dashboard</a>
        <form method="post" action="/cloud/sheets/signout">
          <button>Sign out of Sheets setup</button>
        </form>
      </header>
      <h1>phooey! finance for Google Sheets</h1>
      <p><a href="/foundation">Explore your Foundation views and edit budget targets</a></p>
      <ErrorNotice error={error || state.error || bank.error} />
      <p>
        <a href="/business/tiller-connect/">
          Connect a copy of an existing Tiller workbook, including Estimated
          Quarterly Tax
        </a>
      </p>
      <Panel title="Your workbook">
        <div className="panel-body form-stack">
          <p>
            Transactions and reports refresh twice daily. Stage changes in the
            Edits tab to review and save them from the Phooey sidebar.
          </p>
          {b?.url && (
            <a href={b.url} target="_blank" rel="noreferrer">
              Open your Google workbook
            </a>
          )}
          {b?.synced && (
            <small>
              Last refreshed {new Date(b.synced * 1000).toLocaleString()}
            </small>
          )}
          <ErrorNotice error={b?.error ?? null} />
          <div className="actions">
            <Button
              busy={busy}
              onClick={() =>
                run(async () => {
                  location.assign(
                    (await cloud<{ url: string }>("connect", true)).url,
                  );
                })
              }
            >
              {b?.connected
                ? "Reconnect Google"
                : "Connect Google and create workbook"}
            </Button>
            {b?.connected && (
              <>
                <Button
                  busy={busy}
                  onClick={() =>
                    run(async () => {
                      await cloud("refresh", true);
                    })
                  }
                >
                  Sync now
                </Button>
                <button
                  disabled={busy}
                  onClick={() => {
                    if (
                      confirm(
                        "Disconnect this workbook and revoke its grant? The file stays in Google Drive.",
                      )
                    )
                      void run(async () => {
                        await cloud("disconnect", true);
                        setCode("");
                      });
                  }}
                >
                  Disconnect
                </button>
              </>
            )}
          </div>
          {b?.url && (
            <>
              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  void run(async () => {
                    await cloud("year", true, { year: Number(year || b.year) });
                    setYear("");
                  });
                }}
              >
                <label>
                  Workbook tax year
                  <input
                    type="number"
                    min="2000"
                    max="2100"
                    value={year || b.year}
                    onChange={(e) => setYear(e.target.value)}
                  />
                </label>
                <button disabled={busy}>Update tax year</button>
              </form>
              <h3>Connect the Sheets sidebar</h3>
              {b.addon ? (
                <a href={b.addon} target="_blank" rel="noreferrer">
                  Install the Phooey add-on
                </a>
              ) : (
                <p>
                  The add-on is available to beta testers through a test
                  deployment. Contact support for access.
                </p>
              )}
              <div className="actions">
                <button
                  disabled={busy}
                  onClick={() =>
                    run(async () =>
                      setCode(
                        (await cloud<{ code: string }>("pair", true)).code,
                      ),
                    )
                  }
                >
                  Get pairing code
                </button>
                {b.paired && (
                  <button
                    disabled={busy}
                    onClick={() =>
                      run(async () => {
                        await cloud("revoke", true);
                        setCode("");
                      })
                    }
                  >
                    Revoke sidebar access
                  </button>
                )}
              </div>
              {code && (
                <p>
                  In the sidebar enter <strong>{location.origin}</strong> and
                  this five-minute code <code>{code}</code>.
                </p>
              )}
            </>
          )}
        </div>
      </Panel>
      <Panel title="CSV and manual accounts">
        <div className="panel-body form-stack">
          <form
            onSubmit={(e) => {
              e.preventDefault();
              void run(async () => {
                await post("/accounts", {
                  name: account,
                  kind: "checking",
                  currency: "USD",
                });
                setAccount("");
              });
            }}
          >
            <label>
              New USD account
              <input
                required
                value={account}
                maxLength={100}
                onChange={(e) => setAccount(e.target.value)}
              />
            </label>
            <button disabled={busy}>Add account</button>
          </form>
          <button disabled={!accounts.length} onClick={() => setCsv(true)}>
            Import CSV
          </button>
          <details>
            <summary>Switch from Tiller</summary>
            <p>
              Import your first bank feed before checking for duplicates. Match
              each source account to the correct currency and map its
              categories.
            </p>
            <TillerImport />
            <button
              disabled={busy}
              onClick={() =>
                run(async () => setMatches(await cloud("duplicates")))
              }
            >
              Check imported rows against bank feeds
            </button>
            {matches && (
              <div>
                <p>
                  {matches.length
                    ? matches.length +
                      " potential pairs (up to 100 shown). Review the original transactions before deleting a duplicate."
                    : "No matching date and amount in this account."}
                </p>
                {matches.map((m) => (
                  <p key={m.imported + m.bank}>
                    {m.posted} · {m.account} · {(m.amount / 100).toFixed(2)}{" "}
                    {m.currency}
                    <br />
                    CSV: {m.imported_description}
                    <br />
                    Bank: {m.bank_description}
                  </p>
                ))}
              </div>
            )}
          </details>
        </div>
      </Panel>
      <Panel title="Categories and quarterly estimates">
        <div className="panel-body form-stack">
          <p>
            Use an excluded tax category for the personal portion of a mixed
            purchase. Configure the filing status, transaction tag, rates, and
            cutoff date for your tax year before reviewing the estimate.
          </p>
          <div className="actions">
            <button
              onClick={() =>
                setPanel(panel === "categories" ? "" : "categories")
              }
            >
              Manage categories
            </button>
            <button onClick={() => setPanel(panel === "tax" ? "" : "tax")}>
              Review tax assumptions
            </button>
          </div>
          {panel === "categories" && <Categories />}
          {panel === "tax" && <Taxes />}
        </div>
      </Panel>
      <BankConnection connection={bank.data?.connection} />
      {csv && <ImportDialog onClose={() => setCsv(false)} />}
    </main>
  );
}
