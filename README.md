Member Offset Application Tool for STAAD.Pro
<div align="center">
Automate Member Offset Assignments in STAAD.Pro

Powerful VBA script to analyze and automatically apply offsets to structural members based on node coordinates and section properties.

Developed by Engr. Lowrence Scott D. Gutierrez
LinkedIn

📊 Project Stats
GitHub Views
GitHub Stars
VBA Version
License

Streamline your structural modeling workflow by automatically detecting member types and applying precise offsets based on node coordinates and section properties.

</div>
✨ What Makes This Tool Special
This VBA macro for STAAD.Pro intelligently reads member and node data, classifies members into beams and columns based on coordinate change patterns, and applies suitable offsets considering member support conditions and section details. It simplifies complex structural adjustments, reduces manual errors, and saves valuable engineering time—especially for projects with numerous members.

Note:
Currently, this code only works for beams.
Columns are assumed to be rectangular or circular sections. If a column is rectangular, its depth should be aligned parallel to the global Z-axis. See the "How to Use" section for detailed instructions on section orientation.

Important:
For steel sections to be properly recognized and processed, the file AISC_STEEL_DATABASE.xlsx must be placed in the same directory as your STAAD file. This external database is used to look up steel section properties.

Looking ahead, I plan to improve the code to:

Seamlessly detect column orientation and support offsets for columns
Apply offsets even when columns are steel sections
Enhance offset calculation accuracy and flexibility
🚀 Key Features
<table> <tr> <td width="50%">
🎯 Automated Member Classification
Detects beams and columns based on coordinate change patterns
Handles single-axis (X, Y, Z) and multi-axis member variations
Supports complex bracing detection
🧰 Section Property Extraction
Parses section labels for RECT, CIRC, STEEL, and custom sections
Retrieves section dimensions directly from STAAD properties
Integrates external steel database lookup for steel sections (Excel-based, must be placed in the same directory as your STAAD file)
🔧 Intelligent Offset Calculation
Applies vertical offsets based on member depth
Calculates horizontal offsets considering support member types and connection conditions
Supports offset application to beams and columns with precise control
🖥️ Robust Error Handling & Logging
Captures connection issues with STAAD.Pro
Reports member detection and offset application status
Provides detailed debug information for troubleshooting
📊 Summary & Reporting
Finalizes with a comprehensive summary dialog
Counts total members processed, offsets applied, and failures
Easy to extend for integration with other automation tools
</td> <td width="50%">
⚙️ Compatibility & Requirements
Requires STAAD.Pro with VBA enabled
Works with STAAD's COM automation interface
Compatible with Windows environments
VBA version 7.0 or later (Excel VBA)
🔧 Prerequisites
STAAD.Pro installed and accessible via COM automation
External steel database AISC_STEEL_DATABASE.xlsx placed in the same directory as your STAAD file
Basic knowledge of VBA and STAAD structural modeling
Optional: External steel database in Excel format for steel sections
📝 Usage
Insert the macro into STAAD's VBA editor
Run Main() to start the offset analysis
Review debug logs and message boxes for results
Verify member offsets in STAAD
📁 How to Use
Step 1: Prepare your Model
Ensure all members are properly modeled with section labels
Save your STAAD model
Enable VBA macros in STAAD
Step 2: Add the VBA Script to STAAD
How to add your VBA code to STAAD's UserTools:

Open STAAD.Pro
Go to Tools > VBA Editor (or press ALT + F11)
In the VBA editor, select Insert > Module
Paste your VBA code (the provided script) into the module
Save the macro, e.g., as MemberOffsetTool.bas
Step 3: Add the Macro to STAAD's UserTools
In STAAD, go to Tools > User Tools > Configure
Click Add and browse to select your saved VBA macro file
Assign a name, e.g., "Member Offset Tool"
Save and close
Step 4: Place the Steel Database
Ensure that the file AISC_STEEL_DATABASE.xlsx is placed in the same directory where your STAAD file is located.
This is necessary for the script to recognize and look up steel section properties.
Step 5: Run the Macro
In STAAD, go to Tools > User Tools > select Member Offset Tool
The macro will run, analyze members, and apply offsets
Watch the debug window for detailed logs
Confirm adjustments in the model
🔍 How the Code Works
Connection to STAAD.Pro:
Establishes COM automation connection with STAAD.
Member List Extraction:
Retrieves a list of all members; attempts to use GetBeamList(), falls back to sequential numbering if necessary.
Member Analysis Loop:
For each member:
Retrieves start/end node IDs
Gets node coordinates
Reads section label and properties
Calculates absolute coordinate changes
Classifies member type (Beam or Column)
For columns, assumes rectangular or circular sections. If rectangular, ensure the depth is aligned parallel to Z-axis (see instructions).
Determines support member info at start/end nodes
Calculates offsets based on member type and support connection
Applies offsets via CreateMemberOffsetSpec() and AssignMemberSpecToBeam()
Offset Application:
Uses CreateMemberOffsetSpec() and AssignMemberSpecToBeam() to assign offsets in STAAD.
Summary & Feedback:
Displays total processed members, offsets applied, and any failures.
🧮 Special Notes & Future Plans
The current implementation only applies offsets to beams. Handling columns with offsets, especially steel columns, will be a future enhancement.
The code assumes column sections are rectangular or circular. For rectangular columns, ensure the section's depth is parallel to the global Z-axis by naming or modeling accordingly.
To facilitate this, name your sections properly (e.g., Rect 0.40x0.25) and verify section orientation in your STAAD model.
In future versions, I aim to:
Seamlessly detect column orientation
Enable offset application to columns, including steel sections
Improve the offset calculation logic for more complex scenarios
🧾 License & Credits
This VBA script is developed and maintained by Engr. Lowrence Scott D. Gutierrez.

License: MIT License

Special thanks to:  

STAAD.Pro Development Team
Structural Engineering Community
