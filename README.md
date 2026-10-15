# GOP Generator — Excel VBA Automation

An Excel VBA-based automation solution designed to streamline the generation and management of Guarantees of Payment (GOPs) within healthcare administration workflows.

## Project Overview

The GOP Generator automates repetitive administrative processes by integrating Microsoft Excel, Microsoft Word, PDF generation, and external currency exchange-rate data.

The solution demonstrates how VBA can be used to build structured business workflows, validate information, implement approval controls, and reduce manual document preparation.

## Key Features

- **Automated GOP generation:** Populates Microsoft Word templates using information stored in Excel and exports completed documents as PDFs.
- **Approval workflow:** Validates authorization codes against predefined approval thresholds.
- **Register management:** Records processed GOP transactions in Excel worksheets.
- **Currency exchange-rate integration:** Retrieves exchange rates from the Frankfurter API.
- **Error handling:** Includes validation and error-handling procedures to help prevent incomplete processing.

## Technologies

- Microsoft Excel VBA
- Microsoft Word Automation
- REST API integration (Frankfurter API)
- PDF export
- Windows file-system operations

## Source Code

The `src/` directory contains three VBA modules:

| Module | Description |
|---|---|
| `Module1.bas` | Core GOP generation, document population, approval validation, and register updates |
| `Module2.bas` | Folder validation and local file-system checks |
| `Module6.bas` | Automated currency exchange-rate retrieval and worksheet updates |

## Notes

This public repository contains VBA source code with demonstration credentials and anonymized configuration values. It does not include patient records or the operational Excel workbook.

The Word template is not currently included.
