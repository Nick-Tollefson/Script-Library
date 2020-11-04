<#
  .SYNOPSIS
   Connects to MS Online and uses the data thats been filled out to genreate an excel phone list users can print out

  .DESCRIPTION
  The Get-Phonelist.ps1 script updates the registry with new data generated
  during the past month and generates a report.

  .PARAMETER Console
  # !TODO
  Outputs phonelist to the console instead of an excel sheet.

  .PARAMETER OutputPath
  # !TODO
  Specifies the name and path for the Excel-based output file. By default,
  Get-Phonelist.ps1 generates a name from the date and time it runs, and
  saves the output in the local directory.

  .INPUTS
  None. You cannot pipe objects to Update-Month.ps1.

  .OUTPUTS
  None. Update-Month.ps1 does not generate any output.

  .EXAMPLE
  C:\PS> .\Update-Month.ps1

  .EXAMPLE
  C:\PS> .\Update-Month.ps1 -Console

  .EXAMPLE
  C:\PS> .\Update-Month.ps1 -outputPath C:\Reports\2009\January.csv

  .AUTHOR

#>

# !TODO Add switch to output as list in the console

# Install Excel Module
if (!(Get-Module "ImportExcel")) {
    Install-Module ImportExcel -scope CurrentUser
}

Write-Verbose "Checking MSOL Connection"

# Connect to MSOl if you are not already connected
try {
    Get-MsolDomain -ErrorAction Stop > $null
}
catch {
    if ($null -eq $cred) { $cred = Get-Credential $O365Adminuser }
    Write-Output "Connecting to Office 365 please find the login window to continue"
    Connect-MsolService
    Write-Verbose "MSOL has been connected"
}

Write-Verbose "Getting User List"

# Get users from MSOl that have licences and Last names, this eliminates shared mailboxs and service accounts
# This also adds alias to each property we are going to use in the phone list so we dont have to rename them later
$user_list = Get-MsolUser | Where-Object IsLicensed | Where-Object LastName |
    Select-Object -Property FirstName, LastName, Title, Office, PhoneNumber, MobilePhone |
    Add-Member -MemberType AliasProperty -Name "First Name" -Value FirstName -PassThru |
    Add-Member -MemberType AliasProperty -Name "Last Name" -Value LastName -PassThru |
    Add-Member -MemberType AliasProperty -Name "Office Phone" -Value PhoneNumber -PassThru |
    Add-Member -MemberType AliasProperty -Name "Cell Phone" -Value MobilePhone -PassThru |
    Sort-Object LastName

# Gets date for output file title
$CurrentMonth = Get-Date -UFormat %m
$CurrentMonth = (Get-Culture).DateTimeFormat.GetMonthName($CurrentMonth)
$CurrentYear = get-date -Format yyyy
$reportTitle = "$CurrentMonth $CurrentYear Phone List"

# Styles Excel Sheet so all coulms fit on one page when printed
$styles = $(
    # Name Style
    New-ExcelStyle -FontSize 11 -Range "A1:B100" -FontName 'Verdana' -Height 19 -Width 13 -VerticalAlignment Center
    # Office Style
    New-ExcelStyle -FontSize 7 -Range "C1:C100" -FontName 'Verdana' -Width 11 -VerticalAlignment Center -HorizontalAlignment Center -WrapText
    # Job Title Style
    New-ExcelStyle -FontSize 7 -Range "D1:D100" -FontName 'Verdana' -Width 18 -VerticalAlignment Center -HorizontalAlignment Center -WrapText
    # Office Phone Number Style
    New-ExcelStyle -FontSize 10 -Range "E1:E100" -VerticalAlignment Center -HorizontalAlignment Center -Width 17 -NumberFormat "(000) 000-0000" -FontName 'Courier New' -Bold
    # Cell Phone Numbers Style
    New-ExcelStyle -FontSize 10 -Range "F1:F100" -VerticalAlignment Center -HorizontalAlignment Center -Width 17 -NumberFormat "(000) 000-0000" -FontName 'Courier New' -Bold
    # Title Style
    New-ExcelStyle -FontSize 15 -Bold -Range "A1:F1" -HorizontalAlignment Center -Merge -FontName 'Calibri' -BorderBottom 11 -BorderColor ([System.Drawing.Color]::FromArgb(68, 84, 106)) -FontColor ([System.Drawing.Color]::FromArgb(68, 84, 106)) -Height 20
    # Label Style
    New-ExcelStyle -FontSize 13 -Bold -Range "A2:F2" -HorizontalAlignment Center -FontName 'Calibri' -BorderBottom 11 -BorderColor ([System.Drawing.Color]::FromArgb(68, 114, 196)) -FontColor ([System.Drawing.Color]::FromArgb(68, 84, 106)) -Height 20 -BackgroundColor White
)

# !TODO Generalize output location and make it a switch option
# Remove the previus file if we already geneated one today
Remove-Item "\\newserver1\Public\PS Results\Phone List\PhoneList $(get-date -f yyyy-MM-dd).xlsx" -ErrorAction SilentlyContinue

Write-Verbose "Exporting Excel File"

$user_list |
    Select-Object -Property "Last Name", "First Name", "Office" , "Title" , "Office Phone", "Cell Phone" |
    Export-Excel "\\newserver1\Public\PS Results\Phone List\PhoneList $(get-date -f yyyy-MM-dd).xlsx" -Title $reportTitle -Style $styles -TableStyle Medium2