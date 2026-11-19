# GOP Generator and tracker

## 1. Intro 

In healthcare administration preparing GOP or other documents can be a routine repetitive task that requires entering same information mostly all of the time increasing the risk of human error, and other administrative errors or delays. 

Therefore, based on the Head of Department GOP tracker was developed with the aim to automate this part of workflow using Microsoft Excel VBA. 

The objective was to with "one-click" create GOP, and store data into a tracker. Plus, enhance the overview over the process by assigning validation codes to the team leaders and display real-time data, so to avoid falling below the approved sum. 

## 2. How does it work?

## 3. Main challenges

## Key Features

### Automated GOP Generation

- Extracts relevant patient, provider, insurance, and financial information from Excel.
- Populates a predefined Microsoft Word template using VBA automation.
- Replaces document placeholders with corresponding worksheet values.
- Generates standardized GOP documents.
- Exports completed documents to PDF.

### Financial Validation and Approval Controls

- Validates required financial information before document generation.
- Applies conditional business rules based on financial thresholds.
- Supports authorization-code validation for transactions requiring additional approval.
- Incorporates controls designed to prevent unauthorized document generation.

### Automated Record Management

- Records GOP-related transactions in structured Excel worksheets.
- Maintains relevant financial and administrative information.
- Reduces repetitive manual data entry.
- Supports consistency between generated documents and internal records.

### File and Folder Management

- Validates required file and folder locations.
- Automates document-saving operations.
- Supports consistent naming and organization of generated files.
- Includes error-handling procedures to reduce workflow interruptions.


## Technologies used

| Technology | Purpose |
|---|---|
| Microsoft Excel | Data storage, user interface, and operational workflow |
| VBA | Automation logic, validation, and business rules |
| Microsoft Word | Template-based GOP document generation |
| COM Automation | Programmatic interaction between Excel and Word |
| PDF Export | Creation of standardized output documents |
| Windows File System | File validation, naming, and document storage |

## Code

You can find the code in the following files: 

| File | Responsibility |
|---|---|
| `Module1.bas` | Main GOP generation and workflow automation |
| `Module2.bas` | File and folder validation |
| `Module3.bas` | Supporting automation functionality |

**Technical Skills Demonstrated**

Provides a detailed overview of the technical competencies applied in the project, including:

- Excel VBA development and modular programming.
- Microsoft Office COM automation.
- Structured data processing and validation.
- Business process automation.
- Conditional approval and authorization logic.
- Automated document generation and PDF export.
- Error handling and file-system management.

## Portfolio Information

This project is part of my software development and business automation portfolio.

It demonstrates how existing business tools can be extended through programming to solve practical operational problems.

The project reflects experience in translating business requirements into technical solutions, designing automated workflows, and integrating multiple Microsoft Office applications.
