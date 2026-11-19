# GOP Generator and tracker

## 1. Intro 

In healthcare administration preparing GOP or other documents can be a routine repetitive task that requires entering same information mostly all of the time increasing the risk of human error, and other administrative errors or delays. 

Therefore, based on the Head of Department GOP tracker was developed with the aim to automate this part of workflow using Microsoft Excel VBA. 

The objective was to with "one-click" create GOP, and store data into a tracker. Plus, enhance the overview over the process by assigning validation codes to the team leaders and display real-time data, so to avoid falling below the approved sum. 

## 2. How does it work?

The file contains two main sheets: FORM and CLIENT(Differently named)

- FORM (see attached picture) contains fields that need to be filled out. The fields are double-checked by a macro to ensure that they are correctly filled in. The coordinator must enter the Date of GOP, Provider, Patient, etc. Once the button 'Create a GOP and Populate Table' is clicked, the GOP is created with the inserted data, and at the same time, the CLIENT sheet is automatically populated. The 'Check Path' button is password-protected and is intended only to verify that the SharePoint folder is correctly connected to the OneDrive folder signed in by the coordinator.

<img width="1877" height="705" alt="Screenshot 2026-11-19 140726" src="https://github.com/user-attachments/assets/19b6496a-4d96-452a-afd2-a8813a8b1fbf" />

- CLIENT sheet contains a table that is automatically populated with GOP data when a GOP is created. The columns of the table are generally named exactly as the corresponding fields. However, from column J onwards, the table is no longer populated automatically, but the data is entered manually by the Finance Department. This helps to reconcile the finances between the two departments.  

## 3. Key Features

Automated GOP Generation: Populates a predefined Microsoft Word template using VBA automation and exports completed PDF.

Financial Validation and Approval Controls: Supports authorization-code validation for transactions requiring additional approval.

Automated Record Management: Reduces repetitive manual data entry and supports consistency between generated documents and internal records.

File and Folder Management: Automates document-saving operations and supports consistent naming and organization of generated files.

## 3. Main challenges

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
