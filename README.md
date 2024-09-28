# Capital Gains Tax calculator for UK

Calculation for UK capital gains on share sales. In Apps Script for Google Sheets. Supports same day/month/pool sales and FX conversion.

Input format is a sheet like this:

| Date       | Price (foreign) | Quantity (negative for sales) | FxRate ABC->GBP |
|------------|-----------------|-------------------------------|-----------------|
| DD/MM/YYYY | 123.45          | 10                            | 0.7923          |

