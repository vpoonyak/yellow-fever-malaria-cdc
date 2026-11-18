# Yellow Fever & Malaria Data Extraction

This repository contains a Python script that extracts Yellow Fever and Malaria Prevention Data from the CDC Yellow Book (2026 edition).

## Overview
The script scrapes country-specific information, including:
- Yellow Fever Vaccine Recommendations
- Malaria Transmission Areas
- Drug Resistance
- Plasmodium Species Prevalence
- Recommended Chemoprophylaxis
- Last Updated Date for Each Country

Once extracted, the data is saved as a CSV file.

## Updating from CDC (keeping what changed)
`ExtractMalariaCDC9.ipynb` made the first extraction (CDC's 23 April 2025
edition). To bring it up to date later, run:

```bash
python3 update_from_cdc.py            # fetches CDC's data feed
python3 update_from_cdc.py feed.json  # or reads a saved copy
```

It rewrites `yf_malaria_by_country.csv` and updates the CDC columns in the
`yf_malaria_for_tableau_countrymatch*.xlsx` tables in place (derived columns
and formatting kept). Nothing is lost:

- `history/changes.csv` gets one row per value whose text or links changed:
  country, column, CDC's date for the new text, the date applied, and the
  previous and new HTML. Markup-only differences are applied but not logged.
- `history/yf_malaria_by_country_until_<date>.csv` is the CSV as it was
  before that run.

A change in `Malaria_Has_Transmission` is reported, never applied (CDC's
feed flags Syria while its page shows N/A). The Tableau workbook (`.twbx`)
keeps its own copy of the data: refresh its extract in Tableau afterwards.

First update, 2026-11-18: Colombia, Costa Rica and Ecuador YF entry rules
(CDC, 9 Jul 2025), Mayotte and Yemen malaria (2 and 11 Jun 2026), Bolivia's
caveat, Colombia's entry category, and CDC's moved map image addresses.

## License
This project is licensed under the MIT License.  
Feel free to modify and use this code.

## Contributions
Pull requests and suggestions are welcome. Feel free to fork this repository and contribute.

## Contact
For any issues or suggestions, open a GitHub issue.