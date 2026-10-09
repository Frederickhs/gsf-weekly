SUCCESS FACTORS WEEKLY TRAINING REPORT
Version 0.1.0

PURPOSE
This application creates a weekly training workbook from a Success Factors Weekly Training Excel report.

HOW TO USE
1. Open "Success Factors Weekly Training Report.exe".
2. Click Browse and select the Weekly Training Excel report.
3. Select a Park.
4. Select a Department.
5. Select one or more Organizations / Venues.
6. Review the Selected list and unique employee count.
7. Click Generate Report.
8. Choose whether to open the completed Excel workbook.

REPORT CONTENTS
- Summary worksheet
- One worksheet for each selected Organization / Venue
- Technician first and last name when available
- SSO / Employee ID
- Training item title
- Days Until Due
- Status: Good, Due Soon, or Overdue

STATUS RULES
- Overdue: Days Until Due is less than 0
- Due Soon: Days Until Due is from 0 through 15
- Good: The employee has no training item due within 15 days or overdue

FOLDERS
- Output: Contains the newest generated workbook
- Archive: Contains previous Output contents in timestamped folders
- Logs: Reserved for application logs
- _internal: Required application files. Do not delete, rename, or move this folder.

IMPORTANT
- Keep the EXE and the _internal folder together.
- The source Weekly Training workbook can be located anywhere the user can access.
- Do not move, rename, or delete the source workbook while the application is loading or generating a report.
- Close generated workbooks before running another report if Windows reports that a file is in use.
- Do not place source Weekly Training reports inside the _internal folder.

UPDATES
For an update, replace the complete application folder with the newly provided version. Copy any reports that must be retained before replacing an older folder.

SUPPORT INFORMATION
When reporting a problem, provide:
- Application version from VERSION.txt
- Source workbook file name
- Selected Park and Department
- Selected Organizations / Venues
- Exact error message or screenshot
