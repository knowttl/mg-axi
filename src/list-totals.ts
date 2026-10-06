// LIST-TOTALS: the one shared helper behind every list command's uniform
// "N of M" total line (AXI principle: pre-computed aggregates).
//
// A list line names rows returned of total available. The total is the one
// Graph already supplies on the fetched pages (@odata.count, captured by the
// shared session when the query carries $count=true); it is never fetched
// with an extra request and never invented. Where the pages carry no usable
// total, the line stays honest: a complete read names its rows, a partial
// read says more is available.
//
// The shape matches az-axi's countLine and vectra-axi's collection counts:
// a `total` field (number when known, null when not) beside a `count`
// string. The noun is always plural ("groups", never "group"); the helper
// never inflects, so "1 groups" reads exactly like az-axi's "1 of N" lines.
import { AxiError } from "axi-sdk-js";

export interface ListTotals {
  /** Server-supplied total, or null when the pages carried no usable total. */
  total: number | null;
  /** Uniform total line, e.g. "3 of 10 groups" or "3 groups shown, more available". */
  count: string;
}

function checkShown(shown: number): void {
  if (!Number.isSafeInteger(shown) || shown < 0) {
    throw new AxiError(`Invalid list row count ${shown}`, "VALIDATION_ERROR", [
      "List totals count the projected rows actually returned",
    ]);
  }
}

function checkTotal(total: number | null | undefined): number | null {
  if (total === undefined) return null;
  if (total === null) return null;
  if (!Number.isSafeInteger(total) || total < 0) {
    throw new AxiError(`Invalid list total ${total}`, "VALIDATION_ERROR", [
      "A list total is the server-supplied @odata.count or null when unknown; never fabricate one",
    ]);
  }
  return total;
}

export function listTotals(
  shown: number,
  total: number | null | undefined,
  noun: string,
  complete: boolean,
): ListTotals {
  checkShown(shown);
  const known = checkTotal(total);
  if (typeof noun !== "string" || !noun) {
    throw new AxiError("Invalid list noun for totals", "VALIDATION_ERROR", [
      "Name the listed collection in the plural, e.g. groups, members, memberships",
    ]);
  }
  // A partial read always names more: the known total stays visible even
  // when the window happens to match it (az-axi countLine parity), while
  // an unknown total says more is available. A complete read collapses a
  // matching total to the row count.
  if (known !== null && (known !== shown || !complete)) return { total: known, count: `${shown} of ${known} ${noun}` };
  if (complete) return { total: known, count: `${shown} ${noun}` };
  return { total: known, count: `${shown} ${noun} shown, more available` };
}
