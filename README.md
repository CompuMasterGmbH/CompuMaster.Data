# CompuMaster.Data
.NET library with common methods for simplified data access, data table arrangements and data IO

## Some of the many features - all .Net data providers
* easy and stable access to data from external data sources (native, OLE-DB, ODBC) 
* execute/query data from database with the most less amount of code lines 
* never forget to close a connection and run into pooling problems
* common methods to query/update data from/on database for all .Net data providers - don't care any more about MS SQL Server, MySql, Oracle, PostgreSql or many other RDMS, don't care any more about accessing the RDMS using native .Net data providers, OLE-DB or ODBC connections
* common methods to write back your locally changed data to your external data source
* execute/query data from database with the most less amount of code lines 

## Some of the many features - CSV
* native access to CSV files
  * read/write CSV files with one line of code
```vb.net
        SourceTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile("c:\temp\input.txt", True)
        CompuMaster.Data.Csv.WriteDataTableToCsvFile("C:\temp\data.csv", SourceTable)
```
  * support column headers enabled or disabled
  * culture dependent or culture independent (especially column separtors, decimal separators)
  * always use correct file encoding (ANSI, UTF-8, UTF-16, WINDOWS-1252, ISO-8859-1 and many others)
  * always create valid CSV data
  * support for multi-line data in your CSV files - even Microsoft Excel will understand it on import
* CSV files can use column separators of fixed column widths
* read/write from/to files, strings or 
* read directly a CSV file from an URL
* ideal for your simple REST web service client/server

## DataTables - your in-memory relational database system
* extract columns and rows using filters, where-clauses and many other technics you already know from your favorite SQL system
* join several DataTables in memory as you do with your favorite SQL system
  * Inner Join
  * Left Join
  * Right Join
  * Full Outer Join
  * Cross Join
  * use 1 or more columns for joining
```vb.net
        Dim NewTable As DataTable = _
              CompuMaster.Data.DataTables.SqlJoinTables( _
              LeftTable, New String() {"ID"}, _
              RightTable, New String() {"PrimaryKeyID"}, _
              CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)
```          
* rearrange columns, sort rows and put them into new DataTables
* convert full DataTables, DataSets or just some DataRows into beautiful plain text or HTML tables with just 1 line of code
```vb.net
        CompuMaster.Data.DataTables.ConvertToPlainTextTable(SourceTable)
        CompuMaster.Data.DataTables.ConvertToHtmlTable(SourceTable)
```

## Some more features - XLS(X)
read directly from .XLS/.XLSX files using XlsReader (may require separate database drivers from Microsoft (Office) installed on your system)
```vb.net
SourceTable = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile("c:\temp\input.txt", "sheet1")
```

## There is still more...
Still not convinced? Download your library now and for free and see the many stuff in the library you need all the days regardless if you're a C#, VB.NET or ... developer

* String/Object/Value/Double checks for DbNull.Value, null/Nothing, String.Empty, Double.NaN
* Remove password part from ConnectionString in case you need to show it to your administrative user
* Query from LDAP directories directly into DataTables
* Create empty Microsoft Excel or Microsoft Access database files for immediate read/write access
