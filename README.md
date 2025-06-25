# Effort-Analysis
Implemented standards in the chargeline, which focus on effective and efficient track on effort 

// workday scripts

let
    Source = Excel.Workbook(File.Contents("C:\Users\yc224f\Documents\SkillCode\New\New Raw Data.xlsx"), null, true),
    Workday_Sheet = Source{[Item="Workday",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Workday_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Status", type text}, {"Shift Date", type date}, {"Bems Id", Int64.Type}, {"Employee Name", type text}, {"Under Rpt Dt", type date}, {"Approv Proc Dt", type date}, {"Approv Proc Tm", type datetime}, {"Time Rptg Nm", type text}, {"Ern Cd", type any}, {"Units Amt", type number}, {"Unit Of Measure", type text}, {"Mgr Bems Id", Int64.Type}, {"Manager Name", type text}, {"Company Cd", type text}, {"Company Nm", type text}, {"Activity ID", type text}, {"Work Group Cd", type text}, {"Cost Center of Worker", type text}, {"Cost Center for Time", type any}, {"Acctg Bus Unit Cd", type text}, {"Acctg Location Cd", type text}, {"Acctg Dept Cd", type text}, {"Country Cd", type text}, {"Country Nm", type text}, {"Payable Status Cd", type text}, {"Payable Status Nm", type text}, {"Pay End Date", type any}, {"Paygroup Cd", type text}, {"Paygroup Nm", type text}, {"Field 1", type any}, {"Field 2", type any}, {"Field 3", type any}, {"Team Cd", type any}, {"Comment", type text}, {"Work Proc Detail", type any}, {"Create Dt", type date}, {"Create Tm", type datetime}, {"Last Approving Name", type text}, {"Last Approving Id", Int64.Type}, {"Last Modify Dt", type date}, {"Last Modify Tm", type datetime}, {"Orig Approving Name", type text}, {"Orig Approving Id", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Bems Id", "Eng Bems Id"}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Renamed Columns", "Manager Name", Splitter.SplitTextByEachDelimiter({"("}, QuoteStyle.Csv, true), {"Manager Name.1", "Manager Name.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Manager Name.1", type text}, {"Manager Name.2", type text}}),
    #"Extracted Text After Delimiter" = Table.TransformColumns(#"Changed Type1", {{"Activity ID", each Text.AfterDelimiter(_, "- "), type text}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Extracted Text After Delimiter",{{"Manager Name.1", "Manager Name"}, {"Comment", "Comments"}, {"Under Rpt Dt", "Timesheet Submitted On"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns1",{"Manager Name.2"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each true),
    #"Renamed Columns2" = Table.RenameColumns(#"Filtered Rows",{{"Employee Name", "Engineer Name"}}),
    #"Trimmed Text" = Table.TransformColumns(#"Renamed Columns2",{{"Engineer Name", Text.Trim, type text}, {"Manager Name", Text.Trim, type text}}),
    #"Filtered Rows1" = Table.SelectRows(#"Trimmed Text", each [Time Rptg Nm] = "Regular Hours" or [Time Rptg Nm] = "Hours Greater than Scheduled for Day (Hidden)" or [Time Rptg Nm] = "TOIL"),
    #"Extracted Date" = Table.TransformColumns(#"Filtered Rows1",{{"Comments", Text.Trim, type text}}),
    #"Filtered Rows2" = Table.SelectRows(#"Extracted Date", each true)
in
    #"Filtered Rows2"

// Beeline Script

let
    Source = Excel.Workbook(File.Contents("C:\Users\yc224f\Documents\SkillCode\New\New Raw Data.xlsx"), null, true),
    Beeline_Sheet = Source{[Item="Beeline",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Beeline_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Status", type text}, {"Bems ID", type text}, {"Employee Name", type text}, {"Supplier", type text}, {"Period", type text}, {"Week Ending Date", type date}, {"Approv Proc Dt", type datetime}, {"Work Date ", type date}, {"Time Rptg Nm", type text}, {"Units Amt", Int64.Type}, {"Unit Of Measure", type text}, {"Mgr Bems Id", Int64.Type}, {"Manager Name", type text}, {"Activity ID", type text}, {"Cost Center of Worker", type text}, {"Acctg Bus Unit Cd", type text}, {"Acct Dept Cd", Int64.Type}, {"Create Date", type datetime}, {"Last Approving Name", type text}, {"Last Approving ID", Int64.Type}, {"Last Modified Date", type datetime}, {"Comments", type text}, {"Daily Comments", type text}}),
    #"Extracted First Characters" = Table.TransformColumns(#"Changed Type", {{"Bems ID", each Text.Start(_, 7), type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Extracted First Characters",{{"Bems ID", "Eng Bems ID"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Eng Bems ID", Int64.Type}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"Work Date ", "Timesheet Submitted On"}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Renamed Columns1", "Manager Name", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Manager Name.1", "Manager Name.2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Manager Name.1", type text}, {"Manager Name.2", type text}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Changed Type2",{{"Manager Name.1", "Manager Last Name"}, {"Manager Name.2", "Manager First Name"}}),
    #"Merged Columns" = Table.CombineColumns(#"Renamed Columns2",{"Manager First Name", "Manager Last Name"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Manager Name"),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Merged Columns", "Employee Name", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Employee Name.1", "Employee Name.2"}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"Employee Name.1", type text}, {"Employee Name.2", type text}}),
    #"Merged Columns1" = Table.CombineColumns(#"Changed Type3",{"Employee Name.2", "Employee Name.1"},Combiner.CombineTextByDelimiter(", ", QuoteStyle.None),"Merged"),
    #"Renamed Columns3" = Table.RenameColumns(#"Merged Columns1",{{"Merged", "Engineer Name"}}),
    #"Trimmed Text" = Table.TransformColumns(#"Renamed Columns3",{{"Manager Name", Text.Trim, type text}, {"Last Approving Name", Text.Trim, type text}}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Trimmed Text",{{"Approv Proc Dt", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type4",{"Comments"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Removed Columns",{{"Daily Comments", "Comments"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns4", each ([Time Rptg Nm] = "Regular Time")),
    #"Extracted Date" = Table.TransformColumns(#"Filtered Rows",{{"Comments", Text.Trim, type text}})
in
    #"Extracted Date"


// Consolidated scripts


let
    Source = Table.Combine({Beeline, Workday}),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Status", "Eng Bems ID", "Engineer Name", "Timesheet Submitted On", "Units Amt", "Unit Of Measure", "Mgr Bems Id", "Manager Name", "Activity ID", "Comments", "Eng Bems Id"}),
    #"Extracted Last Characters1" = Table.TransformColumns(#"Removed Other Columns", {{"Activity ID", each Text.End(_, 8), type text}}),
    #"Inserted Text Between Delimiters" = Table.AddColumn(#"Extracted Last Characters1", "Text Between Delimiters", each Text.BetweenDelimiters([Comments], "-", "-"), type text),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Text Between Delimiters",{{"Text Between Delimiters", "Task ID"}}),
    #"Inserted First Characters" = Table.AddColumn(#"Renamed Columns", "First Characters", each Text.Start([Comments], 2), type text),
    #"Renamed Columns1" = Table.RenameColumns(#"Inserted First Characters",{{"First Characters", "Year"}}),
    #"Inserted Last Characters" = Table.AddColumn(#"Renamed Columns1", "Last Characters", each Text.End([Comments], 4), type text),
    #"Renamed Columns2" = Table.RenameColumns(#"Inserted Last Characters",{{"Last Characters", "Job Code"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns2", "Description", each if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "TLEM" then "DE / DAE Mentor connect"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "PMEE" then "DE / DAE Mentee connect"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "KSGI" then "KSS (Beyond SOW) - Presenter"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "KSAT" then "KSS (Beyond SOW) - Attendee"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "WSAT" then "Training / Workshop (External Vendor)"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "CRAT" then "Cross Training - Attended"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "CRGI" then "Cross Training - Given"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "STUN" then "Structures University Training"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "APST" then "Applicant Strengthening Sessions"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "COEP" then "COP, COE Trainings"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "CS" and [Job Code] = "TECP" then "Technical training / Pathways"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "OJ" and [Job Code] = "NTGI" then "New joiners Training - Given"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "OJ" and [Job Code] = "NTAT" then "New joiners Training - Attended"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "OJ" and [Job Code] = "OJGI" then "OJT-Given (Job Rotation)"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "OJ" and [Job Code] = "OJAT" then "OJT-Attended (Job Rotation)"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "ONOA" then "Participating in Innovation Campaign"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "SBID" then "ID Submission"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "OGMT" then "Org Initaitive Meetings"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "AUTO" then "Process  / Production Automation"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "MI++" then "MBE Implementation"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "HPIS" then "High Performance Jouney Sessions"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "VSPI" then "Process improvement"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "VSMM" then "VSM - Implementation"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "EQ" and [Job Code] = "COIM" then "Working on / Deployment of CI"

else if [Activity ID] ="A6001000" and [Task ID] = "EC" and [Job Code] = "ECSS" then "External Conferences / Events"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "QA" and [Job Code] = "ASDQ" then "AS9100D - Documentation / Meeting"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "QA" and [Job Code] = "LEAN" then "Lean - Documentation / Meeting"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "QA" and [Job Code] = "AUPR" then "Audit Preparation"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "QA" and [Job Code] = "INTA" then "Internal Audit - Auditor / Auditee"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "QA" and [Job Code] = "ENTA" then "External Audit - Auditor / Auditee"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "IS" and [Job Code] = "INRS" then "Interview - Resume Screening"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "IS" and [Job Code] = "INTS" then "Interview - Technical interview support"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "TM" and [Job Code] = "AEMS" then "All employees meet"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "TM" and [Job Code] = "SHPT" then "Presenting / Preparing for Stakeholder visit"

else if [Activity ID] ="A6001000" and [Task ID] = "SY" and [Job Code] = "RTTW" then "Ready to take work"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "BC" and [Job Code] = "ESTM" then "New Project / Phase - Estimation"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "BC" and [Job Code] = "CATG" then "New Project / Phase - Categorization"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "BC" and [Job Code] = "PILT" then "New Project / Phase - Pilot Support"

else if [Activity ID] ="A6001000" and [Task ID] = "TR" and [Job Code] = "MDTR" then "Mandatory Trainings"

else if [Activity ID] ="A6001000" and [Task ID] = "DP" and [Job Code] = "ACDP" then "Author / Co-author / Implement  DPs (BDS)"

else if 
	[Activity ID] ="A6001000" and [Task ID] = "SD" and [Job Code] = "HDDT" then "Hardware Downtime"
else if 
	[Activity ID] ="A6001000" and [Task ID] = "SD" and [Job Code] = "SWDT" then "Software Downtime"


else if 
	[Year] = "YY" or [Team Code] = "TC" then "Comments Discrepancy"
else if  
	[Comments] = null then "Blank"

else if
	[Activity ID] <> "A6001000" and [Task ID] = "NX" and [Job Code] = "TKSS" then "Knowledge Sharing Session on a Particular Technical Topic realted to SOW"
else if
	[Activity ID] <> "A6001000" and [Task ID] = "NX" and [Job Code] = "WTCA" then "Weekly / Monthly Tag Up calls with Leads & completing MOMs"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "PSRP" then "Project Status Reporting ( 4 Square, Staff Notes…)"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "QMSW" then "Intreaction with QMS team for the overall quality compliance process, audits"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "LLBP" then "Lessons Learnt & Best Practices"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "PPLN" then "Production planning / process planning, Scheduling, SharePoint creation for new phase"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "RCCA" then "RCCA, Error categorization, pareto analysis"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "5SAV" then "5S related activities"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "AS91" then "To create necessary documents for AS9100 audit"
else if not 
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "PRAD" then "Create Production Trackers, Progress report templates, KPI Templates for Process adherence"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "ARIO" then "RIO Meetings/ management  & RIO Efforts"
else if not 
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "PRIM" then "Process improvement – Time/Cost/Parts"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "SWIN" then "Software installation for new joinee / internally swapped engineers"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "SWIS" then "NX issues/ TCE issues/ VM-Horizon issues/ Redars issues / IT related issues"
else if not
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "DPMS" then "Quality trackers, checklists, Defect metrics, DPM's (CAPA action discussions, DPM Efforts)"
else if not 
	[Activity ID] ="A6001000" and [Task ID] = "NX" and [Job Code] = "APSA" then "Additional Production Support Activities due to Low bandwidth / Awaiting Input"

else if not 
	[Activity ID] ="A6001000" and [Task ID] = "RX" and [Job Code] = "PRIM" then "Process Improvement"
else if not 
	[Activity ID] ="A6001000" and [Task ID] = "RX" and [Job Code] = "PRLE" then "Lessons Learnt / Knowledge development"
else if not 
	[Activity ID] ="A6001000" and [Task ID] = "RX" and [Job Code] = "QMAS" then "QMS Activities"
else if not 
	[Activity ID] ="A6001000" and [Task ID] = "RX" and [Job Code] = "PSUP" then "Project Support Activities"

else if
	[Activity ID] <> "A6001000" and [Task ID] = "DX" and [Job Code] <> {"TKSS","WTCA","PSRP","QMSW","LLBP","PPLN","RCCA","5SAV","AS91","PRAD","ARIO","PRIM","SWIN","SWIS","DPMS","APSA","PRIM","PRLE","QMAS","PSUP"} then "Design Practices"

else "Comments Discrepancy"),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Year", type number}}),
    #"Added Custom1" = Table.AddColumn(#"Changed Type", "Category", each if [Activity ID] ="A6001000" then "Overhead"
else if [Task ID] ="ED" or [Task ID] ="AX"  or  [Task ID] = "EX" or  [Task ID] = "FX" or [Task ID] = "GX" or [Task ID] = "IX" or 
[Task ID] = "PX" or [Task ID] = "QX" or [Task ID] = "SX" or [Task ID] = "TX" or [Task ID] = "UX" then "Production"
else if [Task ID] = "NX" or [Task ID] = "DX" or [Task ID] = "RX" then "Production support"
else if [Task ID] = null then "Blank"
else "Production"),
    #"Inserted Text Range" = Table.AddColumn(#"Added Custom1", "Text Range", each Text.Middle([Comments], 2, 2), type text),
    #"Renamed Columns3" = Table.RenameColumns(#"Inserted Text Range",{{"Text Range", "Team Code"}}),
    #"Extracted Last Characters" = Table.TransformColumns(#"Renamed Columns3", {{"Activity ID", each Text.End(_, 8), type text}}),
    #"Added Custom2" = Table.AddColumn(#"Extracted Last Characters", "Team", each if [Team Code] = "FA" or [Team Code] ="FT" or [Team Code] ="LG" or [Team Code] ="MQ" or [Team Code] ="BD" or [Team Code] ="GE" or [Team Code] ="CL" or [Team Code] ="KC" or [Team Code] ="BG" or [Team Code] ="BI" or [Team Code] ="SP" then "BDS"
else "N/A"),
    #"Added Custom3" = Table.AddColumn(#"Added Custom2", "Project", each if [Team Code] = "FA" then "F-15 | Aft Center Fuselage" else if [Team Code] = "FT" then "F-15 | Fuel Tank" else if [Team Code] = "LG" then "F-15 | Main Landing Gear" else if [Team Code] = "MQ" then "MQ-25 | Fuel Tubing" else if [Team Code] = "BD" then "PST | BDS" else if [Team Code] = "GE" then "Govt Engg" else if [Team Code] = "CL" then "C-17 | Landing Gear" else if [Team Code] = "KC" then "KC-135 | Fan Duct" else if [Team Code] = "BG" then "PST | BGS" else if [Team Code] = "BI" then "PST | BGS Inventory Management" else if [Team Code] = "SP" then "Standard Parts" else if [Team Code] = "EA" then "EASA Payloads" else if [Team Code] = "EC" then "Electrical Connectivity" else if [Team Code] = "EI" then "Electrical Interiors" else if [Team Code] = "FD" then "FAA 787 Payloads" else if [Team Code] = "FL" then "FAA Legacy payloads" else if [Team Code] = "GA" then "Green Aircraft" else if [Team Code] = "IS" then "Interior Stress" else if [Team Code] = "MD" then "Mech | Design" else if [Team Code] = "MS" then "Mech | Stress" else if [Team Code] = "MR" then "Mech | Structures" else if [Team Code] = "SD" then "Interior Stress" else if [Team Code] = "PS" then "Mech | Design" else if [Team Code] = "SS" then "Mech | Stress" else if [Team Code] = "BM" then "Beverage Maker" else if [Team Code] = "" then "Blank" else "Comments Discrepancy"),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Custom3",{"Manager Name", "Team", "Project", "Eng Bems ID", "Engineer Name", "Timesheet Submitted On", "Status", "Units Amt", "Unit Of Measure", "Mgr Bems Id", "Activity ID", "Comments", "Eng Bems Id", "Task ID", "Year", "Job Code", "Description", "Category", "Team Code"}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Reordered Columns1", "Timesheet Submitted On", "Timesheet Submitted On - Copy"),
    #"Renamed Columns4" = Table.RenameColumns(#"Duplicated Column",{{"Timesheet Submitted On - Copy", "Finance Approved On"}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns4",{"Manager Name", "Team", "Project", "Eng Bems ID", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description", "Status", "Units Amt", "Unit Of Measure", "Mgr Bems Id", "Eng Bems Id", "Task ID", "Year", "Job Code", "Team Code"}),
    #"Renamed Columns5" = Table.RenameColumns(#"Reordered Columns2",{{"Units Amt", "Efforts"}}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Renamed Columns5",{"Mgr Bems Id", "Manager Name", "Team", "Project", "Eng Bems ID", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description", "Efforts", "Status", "Unit Of Measure", "Eng Bems Id", "Task ID", "Year", "Job Code", "Team Code"}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Reordered Columns3",{"Mgr Bems Id", "Manager Name", "Team", "Project", "Eng Bems Id", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description", "Efforts", "Status", "Unit Of Measure", "Task ID", "Year", "Job Code", "Team Code"}),
    #"Renamed Columns6" = Table.RenameColumns(#"Reordered Columns4",{{"Efforts", "Efforts (Hrs)"}}),
    #"Inserted Text Range1" = Table.AddColumn(#"Renamed Columns6", "Text Range", each Text.Middle([Comments], 4, 2), type text),
    #"Renamed Columns7" = Table.RenameColumns(#"Inserted Text Range1",{{"Text Range", "Category Code"}}),
    #"Added Custom4" = Table.AddColumn(#"Renamed Columns7", "Overhead Buckets", each if [Activity ID]="A6001000" and [Category Code]="CS" then "Capabilities & Skill Development"
else if [Activity ID]="A6001000" and [Category Code]="OJ" then "OJT"
else if [Activity ID]="A6001000" and [Category Code]="EQ" then "Enablers & QSD"
else if [Activity ID]="A6001000" and [Category Code]="EC" then "External Conferences, Seminars, Events"
else if [Activity ID]="A6001000" and [Category Code]="QA" then "QMS, AS9100D, Internal Audit"
else if [Activity ID]="A6001000" and [Category Code]="IS" then "Hiring & Interview support"
else if [Activity ID]="A6001000" and [Category Code]="TM" then "Team meetings, OPD, Staff, etc.,"
else if [Activity ID]="A6001000" and [Category Code]="SY" then "Standby"
else if [Activity ID]="A6001000" and [Category Code]="BC" then "Backcharging activities
(If Activity IDs is not assigned currently and will be backcharged to respective ID later)"
else if [Activity ID]="A6001000" and [Category Code]="TR" then "Trainings"
else if [Activity ID]="A6001000" and [Category Code]="DP" then "Design Practice"
else if [Activity ID]="A6001000" and [Category Code]="SD" then "System Downtime"

else ""),
    #"Added Custom9" = Table.AddColumn(#"Added Custom4", "Custom.1", each let
    val = [Comments],
    isMatch =
        Text.Length(val) = 14 and
        Text.Range(val, 0, 2) = Text.Select(Text.Range(val, 0, 2), {"2","3","4","5","6","7","8","9"}) and
        Text.Range(val, 2, 4) = Text.Select(Text.Range(val, 2, 4), {"A".."Z"})
        and
        Text.Range(val, 6, 1) = "-"
        and
        Text.Range(val, 7, 2) = Text.Select(Text.Range(val, 7, 2), {"A".."Z"})
        and
        Text.Range(val, 9, 1) = "-"
        and
        Text.Range(val, 10, 4) = Text.Select(Text.Range(val, 10, 4), {"A".."Z"} & {"0","1","2","3","4","5","6","7","8","9"})

in
    isMatch),
    #"Renamed Columns8" = Table.RenameColumns(#"Added Custom9",{{"Custom.1", "Valid Comments"}}),
    #"Added Custom5" = Table.AddColumn(#"Renamed Columns8", "Description(s)", each let
    val = [Comments],
    result =
        if val = null or val = "" or val = " " then 
            "Blank"
        else if Text.Length(val) <> 14 then 
            "Comments Discrepancy"
        else if Text.Range(val, 0, 2) <> Text.Select(Text.Range(val, 0, 2), {"2".."9"}) then 
            "Comments Discrepancy"  // Check: first 2 chars must be 0 or 1
        else if Text.Range(val, 2, 2) <> Text.Select(Text.Range(val, 2, 2), {"A".."Z"}) then 
            "Comments Discrepancy"  // Check: next 2 chars must be uppercase A-Z
        else if Text.Range(val, 4, 2) <> Text.Select(Text.Range(val, 4, 2), {"A".."Z"} & {"0".."9"}) then 
            "Comments Discrepancy"  // Check: next 2 chars must be alphanumeric
		
        else if Text.Middle(val, 6, 1) <> "-" then 
            "Comments Discrepancy"  // Check: dash at position 6
        else if Text.Range(val, 7, 2) <> Text.Select(Text.Range(val, 7, 2), {"A".."Z"}) then 
            "Comments Discrepancy"  // Check: 2 uppercase letters after dash
        else if Text.Middle(val, 9, 1) <> "-" then 
            "Comments Discrepancy"  // Check: dash at position 9
        else if Text.Range(val, 10, 4) <> Text.Select(Text.Range(val, 10, 4), {"A".."Z"} & {"0".."9"}) then 
            "Comments Discrepancy"  // Check: last 4 chars alphanumeric
        else if 
[Comments] = null or [Comments] = "blank" or [Comments] = " " or  [Comments] = "" then "Blank"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "TLEM" then "DE / DAE Mentor connect"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "PMEE" then "DE / DAE Mentee connect"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "KSGI" then "KSS (Beyond SOW) - Presenter"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "KSAT" then "KSS (Beyond SOW) - Attendee"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "WSAT" then "Training / Workshop (External Vendor)"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "CRAT" then "Cross Training - Attended"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "CRGI" then "Cross Training - Given"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "STUN" then "Structures University Training"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "APST" then "Applicant Strengthening Sessions"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "COEP" then "COP, COE Trainings"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "CS" and [Job Code] = "TECP" then "Technical training / Pathways"

else if 
	[Activity ID] ="A6001000" and [Category Code] = "OJ" and [Job Code] = "NTGI" then "New joiners Training - Given"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "OJ" and [Job Code] = "NTAT" then "New joiners Training - Attended"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "OJ" and [Job Code] = "OJGI" then "OJT-Given (Job Rotation)"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "OJ" and [Job Code] = "OJAT" then "OJT-Attended (Job Rotation)"

else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "ONOA" then "Participating in Innovation Campaign"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "SBID" then "ID Submission"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "OGMT" then "Org Initaitive Meetings"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "AUTO" then "Process  / Production Automation"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "HPIS" then "High Performance Jouney Sessions"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "VSPI" then "Process improvement"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "VSMM" then "VSM - Implementation"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "EQ" and [Job Code] = "COIM" then "Working on / Deployment of CI"

else if [Activity ID] ="A6001000" and [Category Code] = "EC" and [Job Code] = "ECSS" then "External Conferences / Events"

else if 
	[Activity ID] ="A6001000" and [Category Code] = "QA" and [Job Code] = "ASDQ" then "AS9100D - Documentation / Meeting"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "QA" and [Job Code] = "LEAN" then "Lean - Documentation / Meeting"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "QA" and [Job Code] = "AUPR" then "Audit Preparation"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "QA" and [Job Code] = "INTA" then "Internal Audit - Auditor / Auditee"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "QA" and [Job Code] = "ENTA" then "External Audit - Auditor / Auditee"

else if 
	[Activity ID] ="A6001000" and [Category Code] = "IS" and [Job Code] = "INRS" then "Interview - Resume Screening"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "IS" and [Job Code] = "INTS" then "Interview - Technical interview support"

else if 
	[Activity ID] ="A6001000" and [Category Code] = "TM" and [Job Code] = "AEMS" then "All employees meet"
else if 
	[Activity ID] ="A6001000" and [Category Code] = "TM" and [Job Code] = "SHPT" then "Presenting / Preparing for Stakeholder visit"

else if
	[Activity ID] ="A6001000" and [Category Code] = "SY" and [Job Code] = "RTTW" then "Ready to take work"

else if
	[Activity ID] ="A6001000" and [Category Code] = "BC" and [Job Code] = "ESTM" then "New Project / Phase - Estimation"
else if
	[Activity ID] ="A6001000" and [Category Code] = "BC" and [Job Code] = "CATG" then "New Project / Phase - Categorization"
else if
	[Activity ID] ="A6001000" and [Category Code] = "BC" and [Job Code] = "PILT" then "New Project / Phase - Pilot Support"

else if
	[Activity ID] ="A6001000" and [Category Code] = "TR" and [Job Code] = "MDTR" then "Mandatory Trainings"

else if
	[Activity ID] ="A6001000" and [Category Code] = "DP" and [Job Code] = "ACDP" then "Author / Co-author / Implement  DPs (BDS)"

else if
	[Activity ID] ="A6001000" and [Category Code] = "SD" and [Job Code] = "HDDT" then "Hardware Downtime"
else if
	[Activity ID] ="A6001000" and [Category Code] = "SD" and [Job Code] = "SWDT" then "Software Downtime"

else if
	[Task ID] = "NX" and [Job Code] = "TKSS" then "KSS (SOW related)"
else if
	[Task ID] = "NX" and [Job Code] = "WTCA" then "Weekly / Monthly Tag Up calls"
else if
	[Task ID] = "NX" and [Job Code] = "PSRP" then "Project Reporting"
else if
	[Task ID] = "NX" and [Job Code] = "QMSW" then "QMS Activities"
else if
	[Task ID] = "NX" and [Job Code] = "LLBP" then "Lessons Learnt & Best Practices"
else if
	[Task ID] = "NX" and [Job Code] = "PPLN" then "Production / Process Planning"
else if
	[Task ID] = "NX" and [Job Code] = "RCCA" then "RCCA"
else if
	[Task ID] = "NX" and [Job Code] = "5SAV" then "5S activities"
else if
	[Task ID] = "NX" and [Job Code] = "AS91" then "AS9100 Documentation"
else if
	[Task ID] = "NX" and [Job Code] = "PRAD" then "Production / Process Trackers & Templates"
else if
	[Task ID] = "NX" and [Job Code] = "ARIO" then "RIO Meetings/ Management"
else if 
	[Task ID] = "NX" and [Job Code] = "PRIM" then "Process improvement"
else if
	[Task ID] = "NX" and [Job Code] = "SWIN" then "Software installation (For New joinees)"
else if
	[Task ID] = "NX" and [Job Code] = "SWIS" then "Network Downtime (During Production)"
else if
	[Task ID] = "NX" and [Job Code] = "DPMS" then "KPI Metrics Report"
else if
	[Task ID] = "NX" and [Job Code] = "APSA" then "Production Support (Due to low bandwidth)"

else if
	[Task ID] = "RX" and [Job Code] = "PRIM" then "Process Improvement"
else if
	[Task ID] = "RX" and [Job Code] = "PRLE" then "Lessons Learnt / Knowledge development"
else if
	[Task ID] = "RX" and [Job Code] = "QMAS" then "QMS Activities"
else if
	[Task ID] = "RX" and [Job Code] = "PSUP" then "Project Support Activities"

else if
	[Task ID] = "DX" and [Job Code] <> {"TKSS","WTCA","PSRP","QMSW","LLBP","PPLN","RCCA","5SAV","AS91","PRAD","ARIO","PRIM","SWIN","SWIS","DPMS","APSA","PRIM","PRLE","QMAS","PSUP"} then "Design Practices"



else if
	[Task ID] = "AX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "Input Study"
else if
	[Task ID] = "PX" and [Job Code] <> "INTG" then "Production"
else if
	[Task ID] = "PX" and [Job Code] = "INTG" then "Integration - MODS Electrical Connectivity"

else if
	[Task ID] = "UX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "SOW Updates"
else if
	[Task ID] = "GX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "Production Guidence (Given)"
else if
	[Task ID] = "QX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "Internal QC"
else if
	[Task ID] = "IX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "Internal Rework"
else if
	[Task ID] = "EX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "External Rework"
else if
	[Task ID] = "ED" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "External Rework (Design Improvement)"
else if
	[Task ID] = "SX" and [Job Code] <> {"INTG","NJTN","NJTR","TRPL","MGCT","ESTM","CATG","PILT"} then "Sign off (External)"
else if
	[Task ID] = "TX" and [Job Code] = "NJTN" then "OJT - Trainee"
else if
	[Task ID] = "TX" and [Job Code] = "NJTR" then "OJT - Trainer"
else if
	[Task ID] = "TX" and [Job Code] = "TRPL" then "Training Plans Preparation"
else if
	[Task ID] = "TX" and [Job Code] = "MGCT" then "Tool / Workshop (SOW related)"
else if
	[Task ID] = "FX" and [Job Code] = "ESTM" then "New Project / Phase - Estimation"
else if
	[Task ID] = "FX" and [Job Code] = "CATG" then "New Project / Phase - Categorization"
else if
	[Task ID] = "FX" and [Job Code] = "PILT" then "New Project / Phase - Pilot Support"
else "Comments Discrepancy"
in
    result),
    #"Reordered Columns5" = Table.ReorderColumns(#"Added Custom5",{"Mgr Bems Id", "Manager Name", "Team", "Project", "Eng Bems Id", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description", "Efforts (Hrs)", "Overhead Buckets", "Status", "Unit Of Measure", "Task ID", "Year", "Job Code", "Team Code", "Category Code"}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Reordered Columns5",{"Mgr Bems Id", "Manager Name", "Team", "Project", "Eng Bems Id", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description(s)", "Efforts (Hrs)", "Overhead Buckets", "Status", "Unit Of Measure", "Task ID", "Year", "Job Code", "Team Code", "Category Code", "Description"}),
    #"Added Custom6" = Table.AddColumn(#"Reordered Columns6", "Team(s)", each if [Team Code] = "EI" or [Team Code] ="EA" or [Team Code]="EC" or [Team Code] ="FP" or [Team Code]="FL" or [Team Code] ="GA" or [Team Code]="IS" or [Team Code] ="MD" or [Team Code]="MS" or [Team Code] ="MR" or [Team Code]="SD" or [Team Code] ="PS" or [Team Code]="SS"
then "BGS-MODS"
else if [Team Code]="FA" or [Team Code] ="FT" or [Team Code]="LG" or [Team Code] ="MQ" or [Team Code]="BD" or [Team Code] ="GE" or [Team Code]="CL" or [Team Code] ="KC" or [Team Code]="BG" or [Team Code] ="BI" or [Team Code]="SP"
then "Govt Engg"
else if [Team Code]="BM" then "External"
else "N/A"),
    #"Reordered Columns7" = Table.ReorderColumns(#"Added Custom6",{"Mgr Bems Id", "Manager Name", "Team(s)", "Project", "Eng Bems Id", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description(s)", "Efforts (Hrs)", "Overhead Buckets", "Status", "Unit Of Measure", "Task ID", "Year", "Job Code", "Team Code", "Category Code", "Description", "Team"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns7",{"Mgr Bems Id", "Status", "Unit Of Measure", "Task ID", "Year", "Job Code", "Team Code", "Category Code", "Description", "Team"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Columns1",{{"Finance Approved On", type date}, {"Timesheet Submitted On", type date}}),
    #"Trimmed Text" = Table.TransformColumns(#"Changed Type1",{{"Comments", Text.Trim, type text}}),
    #"Added Custom7" = Table.AddColumn(#"Trimmed Text", "Engr Bems ID", each if [Eng Bems Id] = null then [Eng Bems ID] else [Eng Bems Id]),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom7",{"Manager Name", "Team(s)", "Project", "Eng Bems ID", "Eng Bems Id", "Engr Bems ID", "Engineer Name", "Timesheet Submitted On", "Finance Approved On", "Activity ID", "Category", "Comments", "Description(s)", "Efforts (Hrs)", "Overhead Buckets"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Eng Bems ID", "Eng Bems Id"})
in
    #"Removed Columns"
