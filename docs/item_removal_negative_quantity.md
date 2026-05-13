# Item Removal Using Negative Quantity

## Overview

Items can be removed from a revision using a **negative quantity** in the Qty column. This is the recommended approach when submitting a revised proposal where a client has requested removal of a previously quoted item. It provides a clear visual signal (the negative number) and produces correct financial calculations.

The alternative is the `REMOVED` scope keyword. Both approaches produce identical lot prices and financial figures — the choice is a matter of preference for how the removal is communicated.

---

## When to Use Each Approach

| Approach | Use when |
|---|---|
| **Negative Qty** | You want to signal removal with a visible number (e.g. `-1 lot`, `-6 ea`). The item stays in the BoM as a clear record. |
| **REMOVED keyword** (Scope column) | You prefer a text label. The item's price and cost cells go blank; only the REMOVED tag is visible. |

Both approaches produce the same lot price, TCDQL, BTCQL, TSPL, TP, and TM at the sheet subtotal level.

---

## How to Apply Negative Quantity

### Full system or lot removal

Set the **Title row** Qty to negative. Leave all sub-items at their original positive quantities.

```
9   PROGRAMMING KIT    -1   lot   ...
  1   Programming USB Cable For R7Ex     6   ea
  2   Programming USB Cable For DM       6   ea
  3   Programming USB Cable For DP       6   ea
```

The lot price and all financial columns (TCDQL, TSPL, TP, TM) automatically go negative, correctly reflecting the credit back to the client.

### Partial removal — one sub-item within a retained lot

Set the **sub-item (Lineitem) row** Qty to negative. Leave the Title row at its original positive quantity.

```
9   PROGRAMMING KIT    1   lot   ...
  1   Programming USB Cable For R7Ex    -6   ea   ← removed
  2   Programming USB Cable For DM       6   ea
  3   Programming USB Cable For DP       6   ea
```

The lot price automatically adjusts downward to reflect only the remaining positive-quantity items. The removed sub-item shows as a line with a negative number — no price is displayed for it (it remains in Lumpsum format like the other sub-items).

### Set-based items

Same rules apply. Negate the Title row to remove the whole set, or negate a sub-item to remove that component from a retained set.

---

## Financial Impact

### Full lot/set removal (negative Title qty)

All financial columns on the Title row go negative:

| Column | Effect |
|---|---|
| G (Subtotal Price) | Negative — the credit amount shown to the client |
| TCDQL | Negative — material cost credit |
| BTCQL | Negative — base cost credit |
| TSPL | Negative — selling price credit |
| TP | Negative — profit credit |
| TM (GM%) | Shows the item's original margin % (informative — signals what was given up) |

The sheet **Subtotal** row aggregates all of these correctly. The project-level GM recalculates to reflect the net position after removal.

### Sub-item removal (negative Lineitem qty)

The lot Title's financial figures reduce proportionally:

| Column | Effect |
|---|---|
| G (Subtotal Price, lot Title) | Reduced — reflects remaining items only |
| TCDQL, BTCQL, TSPL, TP | Reduced accordingly |
| TM (GM%) | Recalculates based on the revised lot figures |

---

## Rules and Constraints

1. **Never negate both the Title and a sub-item within the same lot.** If the Title is already negative (whole lot removed), negating a sub-item within it is meaningless — the sub-item is already excluded by the Title's negative quantity. The result is undefined and should be avoided.

2. **Sub-item negative qty is equivalent to REMOVED for financial purposes.** The lot price, cost, and selling price figures are identical whether a sub-item is marked `REMOVED` or set to a negative quantity. The only difference is visual.

3. **Positive sub-item quantities are not affected.** When you negate a sub-item, all other sub-items in the same lot continue to contribute normally to the lot price.

4. **The lot price cannot go negative through sub-item negation.** Because negative sub-items are excluded from the lot price calculation, the displayed lot price always reflects the sum of the remaining positive-quantity items only. (Note: negating the *Title* row of a lot *can* produce a negative lot total — this is intentional and represents a credit.)

---

## Gross Margin on Removed Items

When the Title row is negated (full removal), the TM column shows the same margin percentage as the original item. This is correct: the percentage reflects the item's inherent profitability, which is unchanged by the removal. The **sheet-level subtotal GM** is the figure to watch — it recalculates to the correct net margin across all remaining and removed items.

When a sub-item is negated (partial removal), the lot's TM recalculates based on the revised lot figures.

---

## Summary of Scope Keywords

For reference, all scope keywords and their effect on pricing:

| Scope | Price shown | Included in totals |
|---|---|---|
| *(blank)* | Yes | Yes |
| `OPTION` | Yes | No |
| `INCLUDED` | No | Yes (cost only) |
| `WAIVED` | No | Yes (cost only) |
| `TBA` | No | No |
| `REMOVED` | No | No |
| **Negative Qty** | No (sub-item) / Negative (Title) | Yes (as credit) |
