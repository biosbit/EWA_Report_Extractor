SAP EWA Action Item Extractor
=============================

Purpose
-------
Windows GUI utility that reads an SAP EarlyWatch Alert report and creates an Excel action tracker.

Supported input
----------------
- DOCX: preferred. The SAP EWA rating icons are read from the DOCX and classified as Red/Yellow/Green.
- PDF: best-effort text parsing. DOCX is recommended when exact rating-icon extraction is required.

Output columns
--------------
Section | EWA Rating | EWA Directives | SAP note | SID | Ticket Reference | Responsible Team | Status | Next Actions

Behavior
--------
- By default, Red and Yellow rated checks are extracted as action items.
- The GUI has an option to include Green-rated checks.
- Section numbers are taken from the numbered EWA headings (the Word AUTONUMLGL heading numbering used by the report).
- EWA Directives are source-derived analysis/recommendation text; the program does not invent recommendations when the report does not provide one.
- SAP Notes are detected from the relevant EWA section and linked to SAP for Me. A SAP Notes sheet lists each referenced note individually with a clickable link.
- SID, Ticket Reference, Responsible Team, Status and Next Actions are left blank for follow-up.
- Status has an Excel dropdown: Open, In Progress, Completed, Deferred, Not Applicable.

Build on Windows
----------------
1. Install Python 3.10+ from python.org and make sure `python` is available in Command Prompt.
2. Put these files in the same folder:
   - ewa_action_extractor.py
   - requirements.txt
   - build_windows.bat
3. Double-click build_windows.bat (or run it from Command Prompt).
4. The executable will be created at:
   dist\EWA_Action_Extractor.exe

Run without building
--------------------
python ewa_action_extractor.py

The program opens a GUI. Select the EWA DOCX/PDF and choose the Excel output location.
