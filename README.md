# Alcohol Pricing Decision System

In 2023, I built and operated the pricing workbooks used to prepare several hundred tequila, mezcal, wine, beer, and spirits SKUs for a new restaurant menu and Toast POS. I was the Director of Operations and Launch Consultant, not the owner or founder.

## The problem

Distributor invoices arrived with inconsistent product descriptions, bottle sizes, and prices. The bar team needed to turn those records into defensible pour and bottle prices without recalculating each SKU from scratch.

The system had to catch description changes, distinguish bottle sizes, compare pour options, handle premium-product exceptions, create clean menu prices, and leave a review path before anything entered Toast.

## What I built

The workflow moved each product through the same decisions:

```text
Distributor invoice
  -> parse description and bottle cost
  -> classify category and bottle size
  -> calculate cost per ounce
  -> compare markup and pour options
  -> review premium and classification exceptions
  -> approve a clean menu price
  -> enter the price in Toast POS
  -> revisit the decision when costs or product mix changed
```

This replaced ad hoc bottle-by-bottle calculation with a common model and a clear approval step. It did not replace operating judgment. Market position, product tier, menu readability, and service goals still affected the final price.

## Core calculations

The original workbooks are not public because they contain historical supplier prices and purchasing terms. The formulas below preserve the operating method.

### Parsing an invoice line

This formula extracted a bottle cost from distributor text:

```excel
=TRIM(MID(D4, FIND("$", D4), FIND(" ", D4, FIND("$", D4)) - FIND("$", D4)))
```

It assumed a space followed the dollar amount. When a distributor changed its text format, the row required review rather than silent acceptance.

### Classifying the product

I used keyword checks to place a product into a working category before costing it:

```excel
=IF(OR(ISNUMBER(SEARCH("Tequila", C4)), ISNUMBER(SEARCH("Tequila", D4))), "Tequila",
 IF(OR(ISNUMBER(SEARCH("Rum", C4)), ISNUMBER(SEARCH("Rum", D4))), "Rum",
 IF(OR(ISNUMBER(SEARCH("Whiskey", C4)), ISNUMBER(SEARCH("Whiskey", D4))), "Whiskey",
 IF(OR(ISNUMBER(SEARCH("Gin", C4)), ISNUMBER(SEARCH("Gin", D4))), "Gin",
 IF(OR(ISNUMBER(SEARCH("Liqueur", C4)), ISNUMBER(SEARCH("Liqueur", D4))), "Liqueur", "")))))
```

A second check handled tequila age statements such as blanco, silver, gold, reposado, anejo, and extra anejo. Blank or unexpected classifications stayed visible for review.

### Converting bottle cost to pour cost

The liquor tab calculated bottle ounces, cost per ounce, and prices for different pours:

```excel
H4: =E4/G4
I4: =F4/H4
K4: =I4*1.5
L4: =K4*4
M4: =I4*2
N4: =M4*4
```

The sequence was:

- Bottle volume divided by milliliters per ounce
- Bottle cost divided by bottle ounces
- Cost for a 1.5-ounce or 2-ounce pour
- A historical 4x menu-price starting point

The multiplier created a comparison point, not an automatic approval.

### Wine and beer logic

The wine tab applied a historical 13 percent load before calculating glass and bottle prices:

```excel
D12: =(C12/25)*1.13
G12: =C12*1.13
H12: =G12*2.5
```

Premium bottles received manual review when the default multiplier produced an unrealistic menu price.

The beer tab calculated loaded unit cost and then applied the selected SKU multiplier:

```excel
E13: =(D13/C13)*1.075
G13: =E13*F13
```

Historical beer multipliers generally ranged from 4x to 5x.

### Toast price matrix

The Toast workbook made the decision easier to compare by showing:

- Landed cost using the historical 15 percent load
- Exact bottle ounces using 29.5735 milliliters per ounce
- Multipliers from 2x to 10x in 0.5 increments
- Prices for 1-ounce, 1.5-ounce, 2-ounce, big-ice, and 2.5-ounce serves
- Bottle service at 30 percent below the by-the-ounce equivalent
- A rounded bottle-service floor and the implied discount

For example:

```excel
Landed cost: =E3+(15%*E3)
Bottle ounces: =M3/29.5735
Selected multiplier: =$O3*P$2
Bottle service: =(AH3*N3)*0.7
Rounded floor: =FLOOR(AM3,100)
Implied discount: =1-(AO3/(AH3*N3))
```

## Controls that mattered

I used manager review rather than an automated test suite. Before a price reached the POS, the reviewer had to:

- Confirm the invoice cost and bottle size
- Inspect blank or unexpected classifications
- Spot-check cost-per-ounce calculations
- Review premium products separately
- Approve a clean guest-facing price
- Revisit the price when supplier cost or product mix changed

These checks were practical for the launch environment, but a reusable product should formalize them with validation rules, an audit log, permissions, and automated tests.

## What the system changed

The workbooks gave operations and bar leadership one repeatable path from an invoice line to an approved POS price. They made exceptions easier to see and reduced the need to rebuild the same calculations for every product.

This repository does not claim a measured profit or margin increase. The system supported pricing decisions, but the available evidence does not isolate its financial impact from sales mix, purchasing, service execution, or other operating changes.

## Evidence and limits

The working spreadsheets, row-level supplier prices, invoice descriptions, vendor accounts, employee data, and product-specific purchasing terms remain private. This public case documents representative historical formulas and the decision process, so readers can assess the method but cannot reproduce a full pricing run.

The assumptions reflect one restaurant in 2023 and may no longer match current taxes, costs, service sizes, or regulations. Keyword classification can fail when invoice text changes. AI assisted with later documentation and confidentiality review; the record does not establish whether AI was used in the original workbook design.

## Related work

- [Restaurant launch program](https://github.com/Samamak1/restaurant-buildout)
- [iTZCALi launch case](https://samamak1.github.io/work/itzcali/)
- [Sama Mushtaq portfolio](https://samamak1.github.io/)
