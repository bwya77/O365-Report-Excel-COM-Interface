# O365-Report-Excel-COM-Interface
[Blog Link](http://thelazyadministrator.com/2018/04/12/office-365-report-using-excel-com-interface-with-powershell/)

# .Description
Generate an Excel workbook report based on infromation gathered from your Office 365 tenant. 

# Report Goals
For my report, I wanted to make sure it could do at least the following:

1. Make a separate WorkSheet for each report
2. Each WorkSheet will have a nice, clean heading
3. Report the following:
  1. Licensed Mailboxes
    1. Users with licenses
    1. Display Name
    
Friendly License Name
If no friendly name found, just use the AccountSkuID
Primary E-Mail Addresses
Alias E-Mail Addresses
Groups
Group Name
Group Type
Distribution
Office 365 Group
Security Group
etc…
Group E-Mail Address (if its got one!)
Shared Mailboxes
Shared Mailbox Name
E-Mail Address
Contacts
Name
E-Mail Address
Public Folders
Name
If It’s Mail-Enabled
E-Mail Address
Domains
Domain Name
If its fully verified
If its default
If it’s default, highlight it
Make the report good looking, and easy to read
Show the data in alphabetical order
Auto-Format Excel rows and columns
