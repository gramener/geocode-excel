[Download Geocode.xlsm](../../raw/master/Geocode.xlsm). This is an Excel file (with a Visual Basic
macro function) that geocodes addresses and reverse geo-codes locations using the
Google's Geocoding API.

Use `=GoogleGeocode(address)` to return the latitude and longitude of an address.

Use `=GoogleReverseGeocode(lat, long)` to return the address of a latitude and longitude.

![GoogleGeocode usage](usage.gif)

This calls the [Google Geocoding API](https://developers.google.com/maps/documentation/geocoding/intro)
without an API key to fetch results.

You are [limited to](https://developers.google.com/maps/documentation/geocoding/usage-limits)
2,500 free requests per day and 50 requests per second. So keep within limits,
*automatic calculations are disabled*. If you drag a formula, select the cells
and press `Ctrl-Q` to run the RefreshSelected macro.

A typical workflow that keeps you within these limits is below:

1. Open [Geocode.xlsm](../../raw/master/Geocode.xlsm). If prompted, Enable Editing and Enable Content.
2. Fill column A with addresses
3. In cell `B2`, type `=GoogleGeocode(A2)`
4. Copy the formula in `B2` down about 30-40 (typically one page)
5. Press `Ctrl-Q` to fetch the values
6. If you see `OVER_QUERY_LIMIT`, try again - eventuallly on a different machine.
7. If you see `ZERO_RESULTS`, try different spellings (e.g. add city, country, or remove parts of the address.)

The result is a `latitude,longitude` string in a single cell. Use
[Data > Text to Columns](https://support.office.com/en-us/article/Split-text-into-different-columns-with-the-Convert-Text-to-Columns-Wizard-30B14928-5550-41F5-97CA-7A3E9C363ED7)
to convert these into 2 column.
