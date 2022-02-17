# CompuMaster.Data
.NET library with common methods for simplified data access, data table arrangements and data IO

[![Github Release](https://img.shields.io/github/release/CompuMasterGmbH/CompuMaster.Data.svg?maxAge=2592000&label=GitHub%20Release)](https://github.com/CompuMasterGmbH/CompuMaster.Data/releases) 
[![NuGet CompuMaster.Data](https://img.shields.io/nuget/v/CompuMaster.Data.svg?maxAge=2592000&label=NuGet%20CM.Data)](https://www.nuget.org/packages/CompuMaster.Data/) 
[![NuGet CompuMaster.Data.Controls](https://img.shields.io/nuget/v/CompuMaster.Data.Controls.svg?maxAge=2592000&label=NuGet%20CM.Data.Controls)](https://www.nuget.org/packages/CompuMaster.Data.Controls) 

## Simple download/installation using NuGet
```powershell
Install-Package CompuMaster.Data
```
respectively
```powershell
Install-Package CompuMaster.Data.Controls
```
Also see: https://www.nuget.org/packages/CompuMaster.Data/

## Some of the many features - all .Net data providers
* easy and stable access to data from external data sources (native, OLE-DB, ODBC) 
* execute/query data from database with the most less amount of code lines: simply **fill a complete DataTable** with **1 (!!) line of code**
```vb.net
        'VB.NET sample:
        Dim MyTable As System.Data.DataTable = CompuMaster.Data.DataQuery.AnyIDataProvider.FillDataTable(
            New System.Data.OleDb.OleDbConnection(ConnectionString),
            "SELECT * FROM table", System.Data.CommandType.Text,
            Nothing,
            CompuMaster.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection,
            "MyTableName")
```
```C#
        // C# sample: 
        System.Data.DataTable myTable = CompuMaster.Data.DataQuery.AnyIDataProvider.FillDataTable(
            New System.Data.OleDb.OleDbConnection(ConnectionString),
            "SELECT * FROM table", System.Data.CommandType.Text,
            null,
            CompuMaster.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection,
            "MyTableName");
```
* execute a **query with several command parameters** within **5 (!!) lines of code**
```vb.net
        Dim MyCmd As New System.Data.SqlClient.SqlCommand("SELECT * FROM table WHERE FN=@FirstName AND LN=@FamilyName", New System.Data.SqlClient.SqlConnection(ConnectionString))
        MyCmd.Parameters.Add("@FirstName", SqlDbType.NVarChar).Value = "John"
        MyCmd.Parameters.Add("@FamilyName", SqlDbType.NVarChar).Value = "O'Regan"
        MyCmd.CommandType = System.Data.CommandType.Text
        MyTable = CompuMaster.Data.DataQuery.AnyIDataProvider.FillDataTable(
            MyCmd,
            CompuMaster.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection,
            "MyTableName")
```
* code reduction = maintenance costs reduction
  * execute/query data from database with the most less amount of code lines seen ever
* code reduction = more stable code
  * forget the Try-Catch-Finally blocks
  * never forget to close a connection and run into pooling problems
* simplified common methods to query/update data from/on database for all .Net data providers
  * don't care any more about MS SQL Server, MySql, Oracle, PostgreSql or many other RDMS
  * don't care any more about accessing the RDMS using native .Net data providers, OLE-DB or ODBC connections
* simplified common methods to write back your locally changed data to your external data source

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
              CompuMaster.Data.DataTables.SqlJoinTables( 
              LeftTable, New String() {"ID"}, 
              RightTable, New String() {"PrimaryKeyID"}, 
              CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)
```          
* rearrange columns, sort rows and put them into new DataTables
* convert full DataTables, DataSets or just some DataRows into beautiful plain text or HTML tables with just 1 line of code
```vb.net
        CompuMaster.Data.DataTables.ConvertToPlainTextTable(SourceTable)
        CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(SourceTable)
        CompuMaster.Data.DataTables.ConvertToHtmlTable(SourceTable)
```

## Some more features - XLS(X)
read directly from .XLS/.XLSX files using XlsReader (may require separate database drivers from Microsoft (Office) installed on your system)
```vb.net
SourceTable = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile( 
    "c:\temp\input.xlsx", 
    "sheet1")
```

## There is still more...
Still not convinced? Download your library now and for free and see the many stuff in the library you need all the days regardless if you're a C#, VB.NET or ... developer

* String/Object/Value/Double checks for DbNull.Value, null/Nothing, String.Empty, Double.NaN
* Remove password part from ConnectionString in case you need to show it to your administrative user
* Query from LDAP directories directly into DataTables
* Create empty Microsoft Excel or Microsoft Access database files for immediate read/write access

## Honors
This library has been developed and maintained by [CompuMaster GmbH](http://www.compumaster.de/) for years.

## References

### CompuMaster.Data.Controls
You may find this library useful for using DataGrids in Windows Forms application with row update support on the foreign data source

```powershell
Install-Package CompuMaster.Data.Controls
```

Also see: https://www.nuget.org/packages/CompuMaster.Data.Controls/
