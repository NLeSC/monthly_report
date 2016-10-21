<#
.Synopsis
File monthly status report e-mails in the appropriate project folder

.DESCRIPTION
For each active project (which has an active project folder), search through e-mails to find a project report.
If a valid project number is found, change the name of the e-mail to reflect project and date,
file the e-mail in the appropriate project folder.

.ASSUMPTIONS
The portfolio folder contains:
- One folder containing both:
-- e-mail reports in which the project number, starting with '27' or '33', is mentioned in the body of the text
-- the current script
- Folder named 'Projects' that contains
-- folders whose name begins with the '27' or '33' project number
Project numbers are unique and all of the same length (8 characters).

.EXAMPLE
Right click on icon of the script and select 'Run with PowerShell'.

.INPUTS
File structure and e-mail reports as described in ASSUMTIONS

.OUTPUTS
e-mail reports renamed and moved to: Projects\27or33...\F- Coordinators\monthly reports\<project nr> report <sent date>.msg
Window with information and warning messages

#>

$outlook = New-Object -comobject outlook.application
ForEach ($project in (Get-ChildItem -Path "..\Projects" -Directory)) {
    If ($project.name.StartsWith("27") -or $project.name.StartsWith("33")) {
#           Loop over the active projects and search for reports
#        Write-Host 'project is ' $project
        $projectnr = $project.name.Substring(0,8)
        $nr_reports = 0
        ForEach ($report in (Get-ChildItem -Path ".\*.msg")) {
#               Loop over the reports in the directory containing this script (type *.msg)
            $msg = $outlook.CreateItemFromTemplate($report)
            If ($msg.body.contains($projectnr)) {
                $nr_reports = $nr_reports + 1
                $mesgdirectory = "..\Projects\" + $project.name + "\F- Coordinators\monthly reports"
                If (!(Test-Path -Path $mesgdirectory)) {
#                       If the 'monthly reports' folder doesn't exist, create it
#                    Write-Host 'Creating directory:' $mesgdirectory
                    New-Item -ItemType directory -Path $mesgdirectory
                }
                $newname = $mesgdirectory + "\" + $projectnr + " report " + $msg.SentOn.ToString("u").Substring(0,10)
#                   Deal with more than one report for a given project
                If ($nr_reports -ne 1) {$newname = $newname + "(" + $nr_reports + ")"}
                $newname = $newname + ".msg"
#                Write-Host 'new report name is' $newname
                If (Test-Path -Path $newname) {
#                       Give warning if report has already been filed
                    Write-Host 'report already exists:' $newname
                } Else {
                    move-item $report $newname
                }
            }
        }
#           Give warnings for unexpected situations
        If ($nr_reports -eq 0) {
            Write-Host 'WARNING No reports found for project' $project
        } ElseIf ($nr_reports -gt 1) {
            Write-Host 'WARNING More than one report found for project' $project
        }
    }
}
ForEach ($report in (Get-ChildItem -Path ".\*.msg")) {
#       Give warning for reports (type *.msg) that are left over
    Write-Host 'WARNING Report:' $report 'does not have an active project folder'
}
Read-Host 'Press Enter to exit...' | Out-Null
