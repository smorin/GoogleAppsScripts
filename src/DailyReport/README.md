This script is code that can be added to a GoogleSheet so that it can be emailed as a daily report.

Add a Sheet called "Control"

//EXAMPLE CONTROL SHEET
```
SUBJECT	SENDEMAIL	EMAILS2SEND	COLUMNS2WRITE	COLUMNS2SORT	DATASHEET	FILTER
Daily Partner Task List Report	Yes	steve@nvent.solutions	Owner	Owner	Task list	{"column":"Status","type":"exclude","values":["Done"]}
		steve.morin@gmail.com	Status	Status		
			Due Date	Due Date		
			Task			
```

You can 
SUBJECT=Subject of the email
SENDEMAIL=Yes or No to enable or disable the report
EMAILS2SEND=list of emails to email the report to
COLUMNS2WRITE=list of column headers to add to the report
COLUMNS2SORT=list of column headers to sort the sheet by
DATASHEET=name of the sheet to pull the data from
FILTER=list of JSON objects to filter records by only exclude is implemented

