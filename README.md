# Alcohol Pricing Decision System

> A governed pricing system converting distributor invoices into consistent cost-per-ounce, markup, and POS menu-price decisions for a high-volume tapas and tequila restaurant.

I built and operated the workbooks in 2023 while serving as **Director of Operations / Launch Consultant**. I was **not the owner or founder** of the restaurant. The system supported pricing decisions across several hundred tequila, mezcal, wine, beer, and spirits SKUs.

## Program brief

| Field | Detail |
|---|---|
| Business challenge | Convert inconsistent distributor invoice lines into repeatable pricing decisions before menu and Toast POS entry. |
| My role | Director of Operations / Launch Consultant; designer and operating user of the pricing workbooks. |
| Users | Operations and bar leadership preparing menus and Toast POS records. |
| Scale | Several hundred SKUs in the historical operating system; the public case documents the decision logic without product-level supplier terms. |
| Status | Historical operating system used during the 2023 restaurant launch and operation; source workbooks are intentionally withheld from the public repository. |
| Primary outcome | Pricing became a governed lookup and review process rather than a bottle-by-bottle calculation from scratch. |

## Repository contents

| File | What it is |
|---|---|
| `README.md` | Recruiter-facing business context, model logic, evidence status, controls, and limitations. |

The working spreadsheets are not published because their row-level inputs include historical supplier purchase prices. The formulas below preserve the operating method while the commercial inputs remain private.

## Operating workflow

```text
Distributor invoice
  -> parse description and bottle cost
  -> classify category and size
  -> calculate cost per ml and cost per ounce
  -> compare markup and pour-size options
  -> select an operating price
  -> enter approved price in Toast POS
  -> review product mix and revise when costs change
```

The model standardized the calculation and made exceptions visible. It did not remove managerial judgment: premium products, market position, clean menu price points, and operational goals still required approval.

## Model detail

The formulas below are historical examples shown as written in the operating workbook. Cell references are retained to make the method auditable; the source values are not published.

### Liquor pricing

Each row represented one invoice item:

```text
category | brand/description | ml | cost | cost/ml | ounces |
cost/oz | 1.5 oz cost | 1.5 oz menu price | 2 oz cost | 2 oz menu price
```

#### Parse bottle cost from the invoice text

```excel
=TRIM(MID(D4, FIND("$", D4), FIND(" ", D4, FIND("$", D4)) - FIND("$", D4)))
```

This formula assumes a space follows the dollar amount and therefore requires exception handling when distributor text changes.

#### Classify the spirit family

```excel
=IF(OR(ISNUMBER(SEARCH("Tequila", C4)), ISNUMBER(SEARCH("Tequila", D4))), "Tequila",
 IF(OR(ISNUMBER(SEARCH("Rum", C4)), ISNUMBER(SEARCH("Rum", D4))), "Rum",
 IF(OR(ISNUMBER(SEARCH("Whiskey", C4)), ISNUMBER(SEARCH("Whiskey", D4))), "Whiskey",
 IF(OR(ISNUMBER(SEARCH("Gin", C4)), ISNUMBER(SEARCH("Gin", D4))), "Gin",
 IF(OR(ISNUMBER(SEARCH("Liqueur", C4)), ISNUMBER(SEARCH("Liqueur", D4))), "Liqueur", "")))))
```

A second classifier handled tequila age statements such as blanco, silver, gold, reposado, anejo, and extra anejo.

#### Convert bottle cost to pour cost

```excel
H4: =SUM(E4/G4)
I4: =SUM(F4/H4)
K4: =SUM(I4*1.5)
L4: =SUM(K4*4)
M4: =SUM(I4*2)
N4: =SUM(M4*4)
```

- `H4`: bottle ml divided by ml per ounce
- `I4`: bottle cost divided by bottle ounces
- `K4` and `M4`: cost of 1.5 oz and 2 oz pours
- `L4` and `N4`: historical 4x menu-price outputs

The unnecessary `SUM()` wrappers are retained because the workbook is an authentic historical operating artifact.

### Wine pricing

The wine tab grossed bottle cost up by a 13% historical load, calculated a glass cost from a 25 oz bottle, and applied a default bottle multiplier.

```excel
D12: =(C12/25)*1.13
G12: =C12*1.13
H12: =G12*2.5
```

- `D12`: per-ounce loaded cost
- `G12`: loaded bottle cost
- `H12`: default 2.5x bottle menu price

Premium bottles were reviewed manually and could receive a lower multiplier when the default produced an unrealistic menu price.

### Beer pricing

```excel
E13: =SUM(D13/C13)*1.075
G13: =SUM(E13*F13)
```

- `E13`: case price divided by unit count, plus the historical purchase-tax load
- `G13`: unit cost multiplied by the selected per-SKU markup

Per-SKU multipliers generally ranged from 4x-5x.

### Toast pricing matrix

The representative Toast workbook calculates:

- Landed cost with the historical 15% load: `=(E3+(15%*E3))`
- Exact bottle ounces: `=(M3/29.5735)`
- Multipliers from 2x-10x in 0.5 increments: `=($O3*P$2)` across 17 columns
- 1 oz, 1.5 oz, 2 oz, big-ice, and 2.5 oz price points
- Bottle service at 30% below the by-the-ounce equivalent: `=(AH3*N3)*0.7`
- A clean bottle-service floor: `=FLOOR(AM3,100)`
- Implied discount tracking: `=(1-(AO3)/(AH3*N3))`

The matrix turned a calculation into a controlled selection: choose the tier, review exceptions, approve a clean price point, and enter it into Toast.

## Controls and validation

The historical workflow relied on manager review rather than automated tests:

- Confirm bottle size and invoice cost before pricing
- Review uncategorized or misclassified descriptions
- Spot-check cost-per-ounce calculations
- Review premium-SKU exceptions separately
- Approve clean menu price points before POS entry
- Compare later sales and product-mix reporting with the pricing decision

These checks should be formalized if the model is reused.

## Evidence and outcome status

| Item | Evidence status |
|---|---|
| Workbook logic | Historical formula patterns documented in this case study; source workbooks withheld. |
| Several hundred SKUs | Leadership-account scope; row-level supplier inputs are not published. |
| Faster, more consistent decisions | Workflow outcome; no timed before-and-after study is claimed. |
| Profit or margin impact | Not independently isolated or claimed by this repository. |

## Role, contributors, and authorship

Sama Mushtaq designed, maintained, and used the pricing workbooks as part of his operations and launch-consulting mandate. Distributor invoices supplied the source descriptions and costs; Toast was the target POS environment. This repository does not claim authorship of vendor data, product brands, or third-party software.

## AI assistance

The public record does not establish AI use in the original 2023 workbook design, so this case makes no unsupported claim about it. AI assisted with the later documentation and confidentiality review of this public presentation; it did not originate the operating inputs, formula logic, or pricing approvals.

## Confidentiality and provenance

- Row-level supplier prices, invoice descriptions, vendor-account details, and product-specific purchasing terms are not published.
- Employee names, roles, wages, time-clock records, and operational source tabs remain private.
- The repository publishes the decision framework and representative formula patterns, not a current or historical vendor price list.
- Product-level examples should be added only after representative values and formula behavior can be verified independently of the source workbooks.

## Limitations

- Manual input and review; not an inventory-management platform.
- Historical 2023 tax and overhead assumptions are context-specific and may no longer apply.
- Because the working spreadsheets are withheld, reviewers can assess the method but cannot reproduce a full pricing run from this repository alone.
- Keyword classification can fail when distributor descriptions change.
- No automated test suite, audit log, user permissions, or database.
- One restaurant context; not validated as a general pricing product.
- Formulas use some unnecessary `SUM()` wrappers and manual premium-tier overrides.
- The system supports decisions; it does not guarantee profitability or regulatory compliance.

## Related

- [Restaurant launch program](https://github.com/Samamak1/restaurant-buildout)
- [iTZCALi launch case](https://samamak1.github.io/work/itzcali/)
- [Sama Mushtaq program leadership portfolio](https://samamak1.github.io/)
