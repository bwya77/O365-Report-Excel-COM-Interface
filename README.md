# O365-Report-Excel-COM-Interface
[Blog Link](http://thelazyadministrator.com/2018/04/12/office-365-report-using-excel-com-interface-with-powershell/)

[YouTube Link](https://www.youtube.com/watch?v=H9WkAttm6Ts)

![Public Folders](http://thelazyadministrator.com/wp-content/uploads/2018/04/PFs.png)

![Groups](http://thelazyadministrator.com/wp-content/uploads/2018/04/groups.png)

# .Description
Generate an Excel workbook report based on infromation gathered from your Office 365 tenant. 

# Report Goals
For my report, I wanted to make sure it could do at least the following:

1. Make a separate WorkSheet for each report
1. Each WorkSheet will have a nice, clean heading
1.Report the following:
	1. Licensed Mailboxes
		1. Users with licenses
		1. Display Name
		1. Friendly License Name
			1. If no friendly name found, just use the AccountSkuID
		1. Primary E-Mail Addresses
		1. Alias E-Mail Addresses
	1. Groups
		1. Group Name
		1. Group Type
			1. Distribution
			1. Office 365 Group
			1. Security Group
			1. etc...
		1. Group E-Mail Address (if its got one!)
	1. Shared Mailboxes
		1. Shared Mailbox Name
		1. E-Mail Address
	1. Contacts
		1. Name
		1. E-Mail Address
	1. Public Folders
		1. Name
		1. If It's Mail-Enabled
		1. E-Mail Address
	1. Domains
		1. Domain Name
		1. If its fully verified
		1. If its default
			1. If it's default, highlight it
1. Make the report good looking, and easy to read
1. Show the data in alphabetical order
1. Auto-Format Excel rows and columns
