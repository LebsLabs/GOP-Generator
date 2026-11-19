# GOP Generator and tracker — Technical Architecture

## 1. System Overview

GOP Generator is a Microsoft Excel VBA automation project designed to automate a departmental administrative process. It is built in Excel worksheet with macros (.xlsm file) plus Word document (input), PDF (output) and Frankfurter API currecny exchange retrieval.  

## 2. System Components

### Module1.bas — GOP Generation

The module1 the main VBA module coordinates the entire process. This module actually could keep the process running even without other two modules. It is responsible for reading the inserted data into the FORM, populate the table in CLIENT sheet and generating a .pdf file as an output which is the GOP document itself, Using a hidden `MAP` worksheet to associate Excel fields with Word document placeholders, Checking approval requirements against configured thresholds, Validating single-use approval codes where required, Updating the relevant Excel register, Handling processing errors and cleanup. The module uses an `APPROVALS` worksheet to support authorization-code validation.

### Module2.bas — Folder Validation

This utility demonstrates folder accessibility checks using VBA.
It includes a demonstration password and a local directory check using Windows environment variables.

### Module3.bas — Exchange-Rate Integration

This module retrieves exchange rates through the public Frankfurter API.
The workflow includes: Reading currency codes from the configured Excel worksheet, Making HTTP requests to retrieve exchange-rate information, Processing the returned data, Updating exchange-rate cells, Recording the latest refresh timestamp, Handling unsuccessful requests and other processing errors.

## 3. Data Flow

The following diagram illustrates the main GOP generation workflow and supporting modules.

```mermaid
flowchart TD
    A["Excel Workbook - GOP input data"] --> B["Module1.bas - Input validation"]
    B --> C{"Approval required?"}
    C -->|Yes| D["APPROVALS worksheet - Validate single-use code"]
    C -->|No| E["MAP worksheet - Field-to-placeholder mapping"]
    D -->|Approved| E
    E --> F["Microsoft Word - Populate template"]
    F --> G["Export GOP as PDF"]
    G --> H["Excel Register - Record GOP transaction"]

    I["Module6.bas - Frankfurter API"] --> J["Exchange-rate cells and refresh timestamp"]
    J -.-> A

    K["Module2.bas - Folder validation"] -.-> L["Local file-system check"]
```

Currency exchange-rate retrieval and folder validation are shown as supporting functions. The diagram represents the conceptual workflow rather than every internal VBA procedure.

## 4. Dependencies

- Microsoft Excel with VBA support.
- Microsoft Word for document generation.
- Windows environment for local file-system operations.
- Internet connectivity for exchange-rate retrieval.
- An appropriately configured Excel workbook and Word template.

## 5. Configuration

The public source code contains demonstration configuration values.

The application depends on specific worksheet names, cell references, mapping tables, and document placeholders.

Users adapting the source code must configure these elements for their own workbook and template.

## 6. Security and Privacy

This repository is intended to demonstrate VBA automation techniques.

- Operational credentials are not included.
- Demonstration passwords must not be treated as secure authentication.
- Patient records and production workbooks are not distributed.
- Internal folder paths have been replaced or anonymized.
- The Word template is not currently included.

## 7. Project Scope

The repository provides VBA source modules for technical reference and portfolio demonstration.

It is not a standalone, ready-to-run application. The original workbook structure and supporting document templates are required for full functionality.
