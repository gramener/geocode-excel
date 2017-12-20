[Download Geocode.xlsm](../../raw/master/Geocode.xlsm). This is an Excel file (with a Visual Basic
macro function) that geocodes addresses and reverse geo-codes locations using the
Google's Geocoding API.

Use `=GoogleGeocode(address)` to return the latitude and longitude of an address.

Use `=GoogleReverseGeocode(lat, long)` to return the address of a latitude and longitude.

![GoogleGeocode usage](usage.gif)

This calls the [Google Geocoding API](https://developers.google.com/maps/documentation/geocoding/intro)
without an API key to fetch results.

You are [limited to](https://developers.google.com/maps/documentation/geocoding/usage-limits)

- 2,500 free requests per day
- 50 requests per second

A typical workflow that keeps you within these limits is below:

1. Open [Geocode.xlsm](../../raw/master/Geocode.xlsm). If prompted, Enable Editing and Enable Content.
2. Fill column A with addresses
3. In cell B2, type `=GoogleGeocode(A2)`
4. Copy the formula in B2 down about 30-40 (typically one page) *and wait*
5. After geocoding, [Paste-Special by Values](https://support.office.com/en-us/article/Keyboard-shortcuts-for-Paste-Special-options-c31b7c9e-69ce-4b60-8c3a-dc5ea10d872c)
  so that the request is not made again
6. If you see a `#VALUE!`, try again with other spellings. Add city, country for context
7. Repeat step 4 onwards

If you hit the API limits, break up the geocoding across different people / machines.

The result is a `latitude,longitude` string in a single cell. Use
[Data > Text to Columns](https://support.office.com/en-us/article/Split-text-into-different-columns-with-the-Convert-Text-to-Columns-Wizard-30B14928-5550-41F5-97CA-7A3E9C363ED7)
to convert these into 2 column.
