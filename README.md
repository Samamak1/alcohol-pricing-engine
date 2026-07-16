# Alcohol Pricing Engine

A cost-to-menu pricing system I built and ran as owner/operator of a tapas & tequila
restaurant in Wichita, KS (2023). Every bottle that came through the door - several hundred SKUs of
tequila, mezcal, wine, and beer - went through this workbook before it got a menu price.

No inventory software, no consultants. Distributor invoices went in one side; POS-ready
pour prices came out the other.

## Files

| File | What it is |
|---|---|
| `Alcohol_Pricing_CLEAN.xlsx` | The master pricing workbook (tabs: `pricing`, `wines`, `beers`), ~3,500 live formulas |
| `Toast_Price_Matrix_Example.xlsx` | A single-invoice pricing run built for Toast POS entry - a full 2x-10x markup matrix per bottle, ~3,600 formulas |

## How the model works

### Liquor (`pricing` tab)

Each row is one bottle from a distributor invoice. The column chain:

```
category | brand/description | ML | cost | cost per ml | ounces | cost per oz | 4x | cost 1.5oz | 4x | cost 2oz | 4x
```

1. **Parse the cost out of the invoice line.** Invoice exports came in as text like
   `"1800 Reposado Tequila (750mL/item) - $25.80"`, so the workbook pulls the dollar
   figure straight out of the description string:

   ```excel
   =TRIM(MID(D4, FIND("$", D4), FIND(" ", D4, FIND("$", D4)) - FIND("$", D4)))
   ```

2. **Auto-categorize the bottle** by keyword search against the description - spirit family:

   ```excel
   =IF(OR(ISNUMBER(SEARCH("Tequila", C4)), ISNUMBER(SEARCH("Tequila", D4))), "Tequila",
    IF(OR(ISNUMBER(SEARCH("Rum", C4)), ISNUMBER(SEARCH("Rum", D4))), "Rum",
    IF(OR(ISNUMBER(SEARCH("Whiskey", C4)), ISNUMBER(SEARCH("Whiskey", D4))), "Whiskey",
    IF(OR(ISNUMBER(SEARCH("Gin", C4)), ISNUMBER(SEARCH("Gin", D4))), "Gin",
    IF(OR(ISNUMBER(SEARCH("Liqueur", C4)), ISNUMBER(SEARCH("Liqueur", D4))), "Liqueur", "")))))
   ```

   A second classifier does the same for tequila age statements (plata/blanco/silver,
   gold, reposado, añejo, extra añejo).

3. **Convert to per-ounce cost.** `ML / 29.57 = ounces`, then `cost / ounces = cost per oz`:

   ```excel
   H4: =SUM(E4/G4)     ' bottle ML / ml-per-oz = ounces in bottle
   I4: =SUM(F4/H4)     ' bottle cost / ounces = cost per ounce
   ```

4. **Markup tiers at 4x.** Pour costs are computed at 1.5 oz (cocktail pour) and 2 oz
   (neat/rocks pour), each marked up 4x - a 25% pour cost:

   ```excel
   K4: =SUM(I4*1.5)    ' cost of a 1.5 oz pour
   L4: =SUM(K4*4)      ' menu price at 4x
   M4: =SUM(I4*2)      ' cost of a 2 oz pour
   N4: =SUM(M4*4)      ' menu price at 4x
   ```

### Wine (`wines` tab)

Bottle cost is grossed up 13% for tax/overhead, glass price is derived from a 25-oz
bottle at an 8-oz pour, and bottle menu price is 2.5x:

```excel
D12: =(C12/25)*1.13   ' per-oz cost including the 13% load
G12: =C12*1.13        ' loaded bottle cost
H12: =G12*2.5         ' bottle menu price at 2.5x
```

(High-end bottles like Dom Pérignon and Laurent-Perrier were dialed back to 2x by hand -
markup discipline bends at the top of the list or the bottles never leave the cellar.)

### Beer (`beers` tab)

Per-unit cost out of the case, plus 7.5% purchase tax, times a per-SKU markup
(4x-5x depending on the beer):

```excel
E13: =SUM(D13/C13)*1.075   ' case price / units, taxed = single cost
G13: =SUM(E13*F13)         ' single cost x markup = menu price
```

### The Toast price matrix

`Toast_Price_Matrix_Example.xlsx` is one full invoice (120+ bottles) prepared for POS
entry. For every bottle it computes:

- Landed cost with 15% purchase tax: `=(E3+(15%*E3))`
- Exact ounces: `=(M3/29.5735)`
- A price matrix at every half-step multiplier from 2x to 10x (`=($O3*P$2)` across 17 columns)
- Toast price points per pour: 1 oz, 1.5 oz (cocktail), 2 oz (neat/rocks), 2 oz big ice (+$2), 2.5 oz (martini)
- Bottle-service pricing: 30% off the by-the-ounce equivalent, floored to a clean $100 step:
  `=(AH3*N3)*0.7` then `=FLOOR(AM3,100)`, with the implied discount tracked as `=(1-(AO3)/(AH3*N3))`

The matrix meant a pricing decision was a lookup, not a calculation: pick the multiplier
that fits the bottle's tier, read across, type into Toast.

## Honest notes

- **The workflow was manual.** One workbook (or tab) per distributor invoice, built as
  invoices arrived - this repo's master file carries the merged result. It worked, and it
  also taught me exactly why restaurants pay for inventory software.
- **Formulas are wrapped in `SUM()` unnecessarily** (`=SUM(E4/G4)` instead of `=E4/G4`).
  Self-taught Excel from the middle of a restaurant opening looks like this.
- **Employee data removed.** The original workbook had a block of employee time-clock
  data (names, roles, wages, clock-in/out times) pasted into spare columns of the pricing
  tab - the spreadsheet was the office. Those columns (originally R-X) have been deleted
  entirely from this copy, along with per-invoice vendor tabs. Every remaining cell has
  been scanned; the only proper names left are wine and spirit brands.
