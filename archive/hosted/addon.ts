import { z } from "zod";
import { Ledger } from "../server/ledger";
import { assert, readJSON } from "../server/util";
import { checkGrant, finishPair, refreshWorkbook } from "./workbook";
import type { HostedEnv } from "./control";
export async function addonRequest(
  env: HostedEnv,
  req: Request,
  owner: string,
) {
  const path = new URL(req.url).pathname;
  assert(req.method === "POST", 405, "Use POST.");
  const file = req.headers.get("X-Phooey-Workbook") ?? "";
  const ledger = new Ledger(env.DB, "sheet:" + owner);
  const body = await readJSON(req, 100_000);
  if (path === "/addon/pair") {
    const { code } = z.object({ code: z.string().length(43) }).parse(body);
    return { token: await finishPair(env, code, file, owner) };
  }
  const token =
    req.headers.get("Authorization")?.match(/^Bearer ([\w-]{43})$/)?.[1] ?? "";
  await checkGrant(env, token, file);
  switch (path) {
    case "/addon/refresh":
      await refreshWorkbook(env);
      return { ok: true };
    case "/addon/options":
      return {
        accounts: await ledger.accounts(),
        categories: await ledger.categories(),
      };
    case "/addon/transaction":
      return ledger.transaction(z.object({ id: z.string() }).parse(body).id);
    case "/addon/save": {
      const data = z
        .object({
          id: z.string().optional(),
          revision: z.number().int().positive().optional(),
          requestKey: z.uuid(),
          transaction: z.unknown(),
        })
        .parse(body);
      assert(
        !data.id === !data.revision,
        400,
        "Existing rows need an ID and revision.",
      );
      return ledger.saveTransaction(data.transaction, data);
    }
    case "/addon/delete": {
      const d = z
        .object({ id: z.string(), revision: z.number().int() })
        .parse(body);
      return ledger.deleteTransaction(d.id, d.revision);
    }
    case "/addon/link": {
      const d = z
        .object({
          first: z.string(),
          second: z.string(),
          revisions: z.tuple([z.number().int(), z.number().int()]),
        })
        .parse(body);
      return ledger.linkTransfer(d.first, d.second, d.revisions);
    }
    case "/addon/unlink": {
      const d = z
        .object({ id: z.string(), revision: z.number().int() })
        .parse(body);
      return ledger.unlinkTransfer(d.id, d.revision);
    }
    default:
      throw new Error("Unknown spreadsheet command");
  }
}
