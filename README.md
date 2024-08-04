The repo provides sample code to automate a use case for sending email for Purchase Order processing.

Basic workflow is as follows:
1. Read emails with PO number in their mail subject.
2. Fetch the details for this specific mail from the excel file. (Details include TO, CC, BCC, Comment)
3. Send the mail accordingly.

The details of the scenario are outlined in Scenario.docx

Timed Logging is also implemented. The logs for each process will be stored in one log file.
The process can be scheduled to be run at specific intervals.
