# TPA Operations Dashboard

A Python dashboard that reads running balances from Excel workbooks, stores current values and historical observations in SQLite, and displays them through a Flask web interface served by Waitress.

**Portfolio edition:** This repository uses anonymized company names and a generator for fictional sample data. Operational workbooks, private paths and real financial records are excluded.

## Why this project was built

The project began with a department manager's request for email notifications when a running balance fell below a threshold. A dashboard was proposed to provide continuous visibility of balances and their movements across several insurance companies.

The intended internal setup is a small laptop connected to two display screens, allowing the department to monitor balances without repeatedly opening individual Excel workbooks. This screen deployment is a planned use case; the public edition demonstrates the application locally. Email alerts are not implemented in this version.
