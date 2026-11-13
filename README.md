# ☕ Coffee Orders Sales Dashboard (Excel)

An interactive Excel dashboard that analyzes about 1,000 coffee orders (2 Jan 2019 to 19 Aug 2022). It shows how sales change over time, which coffee types and countries perform best, and who the top customers are.
## Workbook Sheets

- **orders**: main table with raw data and calculated columns
- **customers**: customer lookup table (name, email, country, loyalty card)
- **products**: product lookup table (coffee, roast, size, price, profit)
- **Pivot Table**: 3 pivot tables that feed the charts
- **Dashboard**: 3 charts and 3 slicers (Size, Roast Type, Loyalty Card)

## What I Did

1. Pulled customer details (name, email, country) with `XLOOKUP`:
   ```excel
   =XLOOKUP(C2, customers!A:A, customers!B:B, , 0)
   ```
2. Pulled product details (coffee type, roast type, size, unit price) with `INDEX + MATCH`:
   ```excel
   =INDEX(products!$A$1:$G$49, MATCH($D2, products!$A:$A, 0), MATCH(I$1, products!$A$1:$G$1, 0))
   ```
3. Used nested `IF` to turn coffee and roast abbreviations into readable names (e.g., `Rob` → Robusta, `L` → Light)
4. Formatted order dates with a 3-letter month to avoid European/American date confusion
5. Formatted size in Kg and unit price and sales in $ using custom number formats
6. Removed duplicate rows
7. Built 3 pivot tables and pivot charts (sales over time, by country, top 5 customers)
8. Grouped the first pivot table by years and months only
9. Hid the field buttons on the charts and added slicers to make the dashboard interactive

10. ## Skills Demonstrated

Excel, XLOOKUP, INDEX/MATCH, nested IF, data cleaning, custom number formats, pivot tables, pivot charts, slicers, dashboard design
