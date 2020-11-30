# RetrieveAndCheckData

This application has been designed with a main idea: connect and see a few rows of any table using 3 drivers: ADODB, SQL CLIENT and OLEDB.

Basically we need to configure the parameters:

## Connectivity

- **$Server** = "xxxxxxx" // Azure SQL Server name OR Managed Instance
- **$Db** = "xxxxxx" // Database Name
- **$User** = "xxxxxx" // User Name
- **$Password** = "xxxxxx" // Password
- **$Provider** = "x"  // The provider type to connect: (1) - OLEDB SQLOLEDB or (2) - OLEDB MSOLEDBSQL or (3) .Net SQL Client or (4) ADO - SQLOLEDB or (5) ADO - MSOLEDBSQL

## Executing this Powershell script, it will be done:
- Connect to the database and run a query that you could specify.
- This tool will retrieve every column in your query, for example, you could use select top(20) * from table_1
      
Enjoy!
