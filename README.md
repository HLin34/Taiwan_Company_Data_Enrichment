# Taiwan Company Data Enrichment

Enrich a list of Taiwanese company names in two stages. The FBFH importer/exporter registration portal supplies business registration IDs and trade information. The GCIS company registry API uses those IDs to look up registration status and capital data.

## Setup

From the repository root, install the dependencies in Python 3.10 or newer.

~~~sh
python -m pip install -r requirements.txt
~~~

The FBFH stage requires Google Chrome, Tkinter, and a desktop session for manual CAPTCHA entry. On Windows or macOS, use `python3` if that is the command for your Python installation.

## 1. Enrich company names with FBFH

Place your workbook at `company_list.xlsx` in the repository root and give it a `公司名稱` column. The first sheet is used by default. The script also accepts `.xls` and CSV files when you set `INPUT_FILE` in `src/fbfh_company_enrichment.py`. To try the input format, set it to the [sample CSV](examples/sample_company_list.csv). The sample uses fictional names, so expect missing matches.

~~~sh
python src/fbfh_company_enrichment.py
~~~

Enter CAPTCHA values through the prompt when requested. The script writes `company_list_fbfh_enriched.xlsx` with sheets `公司基本資料` (company details), `歷年實績級距` (historical trade bands), and `執行摘要` (run summary). It preserves input columns, saves after each company, and skips successful lookups when rerun against the same output. Review ambiguous matches manually and use a fresh output filename for another list.

## 2. Enrich registration data with GCIS

~~~sh
python src/company_registry_api_enrichment.py
~~~

The API stage reads `company_list_fbfh_enriched.xlsx` and looks up the `統一編號` (business registration ID) from `公司基本資料`. It writes `company_list_fbfh_enriched_api.xlsx` and adds fields such as company status, representative, registered and paid-in capital, registration address, and registration dates where returned. Rows without valid IDs cannot be looked up. Keep the API input path aligned with the FBFH output path if you customize either script.

Keep real company lists, output workbooks, and screenshots local. Matches and available fields depend on the source records. Review the data providers' terms before collecting or redistributing results.

## License

The source code is available under the [MIT License](LICENSE). Data obtained from external services remains subject to the terms of its source.
