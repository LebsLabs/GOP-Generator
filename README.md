# GOP Generator — Excel VBA Automation

An Excel VBA-based automation solution designed to generate, manage, and track Guarantees of Payment (GOPs) within healthcare administration workflows.

## Why?
<div align="justify">
The department responsible for issuing Guarantees of Payment to healthcare clinics previously relied on a largely manual process. Each GOP required staff to prepare the document individually, enter its details into an Excel register, and manually track subsequent changes, including whether the GOP had been closed and how much of the approved amount remained available. This approach involved repetitive administrative work, duplicated data entry, and the risk of inconsistencies between documents and registers. The objective was to consolidate these activities into a single automated workflow, minimizing discrepancies between generated GOP documents and their corresponding register entries while providing an all-in-one tool for GOP management. The GOP Generator was developed to enable staff to generate documents, update records, and maintain financial tracking through a streamlined process. An additional requirement was to introduce an approval mechanism, allowing designated department managers to authorize GOPs exceeding predefined financial thresholds using approval codes.
</div>


**The objective was to consolidate these activities into a single automated workflow with the aim to reduce the inconsistency between the GOP and data entered into the table, and all-in-one tool for GOPs.**

The GOP Generator was developed to enable staff to generate documents, update records, and maintain financial tracking through a streamlined process. An additional requirement was to introduce an approval mechanism, allowing designated department managers to authorize GOPs exceeding predefined financial thresholds using approval codes.

## Project Overview

The GOP Generator automates repetitive administrative processes by integrating Microsoft Excel, Microsoft Word, PDF generation, and external currency exchange-rate data. The solution demonstrates how VBA can be used to build structured business workflows, validate information, implement approval controls, and reduce manual document preparation.

## Key Features

| Feature | Description |
|---|---|
| **Automated GOP Generation** | Populates Word templates using information from Excel and exports as PDFs. |
| **Approval Workflow** | Validates authorization codes against predefined financial thresholds. |
| **Register Management** | Records processed GOP transactions in Excel worksheets. |
| **Financial Tracking** | Supports the monitoring of GOP statuses and remaining financial balances. |
| **Currency Exchange-Rate Integration** | Retrieves exchange rates from the Frankfurter API. |
| **Error Handling** | Includes validation and error-handling procedures to help prevent incomplete processing. |

## Technologies

- Microsoft Excel VBA
- Microsoft Word Automation (COM)
- REST API integration (Frankfurter API)
- PDF generation and export
- Windows file-system operations
- Microsoft SharePoint for file storage and collaboration

## Source Code

The `src/` directory contains three VBA modules:

| Module | Description |
|---|---|
| `Module1.bas` | Core GOP generation, document population, approval validation, and register updates |
| `Module2.bas` | Folder validation and local file-system checks |
| `Module6.bas` | Automated currency exchange-rate retrieval and worksheet updates |

## Notes

This public repository contains anonymized VBA source code intended for technical review and portfolio demonstration. The operational Excel workbook, Word template, production credentials, and confidential business records are not included. As a result, the repository is not intended to function as a standalone, ready-to-run application.

## Technical Documentation

For a more detailed understanding of the project, see:

- [System Architecture](docs/ARCHITECTURE.md) — Overview of the application components, VBA modules, data flow, and dependencies.
- [Technical Skills Demonstrated](docs/TECHNICAL_SKILLS.md) — Programming techniques, Microsoft Office automation, API integration, and software development practices demonstrated by the project.

## Portfolio Information

**Developer:** [LebsLabs](https://github.com/LebsLabs)

**Project Type:** Business Process Automation in VBA

**Status:** Portfolio demonstration

This project illustrates the application of VBA programming to a real-world administrative challenge, combining document automation, financial tracking, approval workflows, and external data integration.
