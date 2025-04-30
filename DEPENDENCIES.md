# Dependencies

## Upgrade of followingDependencies ON HOLD

### Definitively not compatible

* CompuMaster.Web.TinyWebServerAdvanced
  * IST: v2021.7.28.100
  * NEU: v2024.11.4.100 -> Unit Tests schlagen fehl mit Test-Host abgestürzt
	* Vorbedingung 1: Test-Platform .NET 8 (.NET 4.8 + 6.0 sind ok, 9.0 nicht getestet)
	* Vorbedingung 2: Tests müssen zusammen im gleichen Test-Lauf in VS ausgeführt werden, damit Test-Host abstürzt
	  * CompuMaster.Test.Data.CsvTest.ReadDataTableFromCsvUrlAtLocalhostWithContentTypeButWithoutCharset
	  * CompuMaster.Test.Data.CsvTest.ReadDataTableFromCsvUrlWithTls12Required
	* Protokollierte Fehlermeldung   
```
NUnit Adapter 4.5.0.0: Test execution started
Running selected tests in D:\bin\Debug\net8.0\CompuMaster.Test.Tools.Data.dll
   NUnit3TestExecutor discovered 2 of 2 NUnit test cases using Current Discovery mode, Non-Explicit run
Unhandled exception. System.Net.HttpListenerException (995): Der E/A-Vorgang wurde wegen eines Threadendes oder einer Anwendungsanforderung abgebrochen.
   at System.Net.HttpListener.GetContext()
   at CompuMaster.Web.TinyWebServerAdvanced.WebServer.<Run>b__20_0(Object o)
   at System.Threading.QueueUserWorkItemCallback.Execute()
   at System.Threading.ThreadPoolWorkQueue.Dispatch()
   at System.Threading.PortableThreadPool.WorkerThread.WorkerThreadStart()
Der aktive Testlauf wurde abgebrochen. Grund: Der Testhostprozess ist abgestürzt. : Unhandled exception. System.Net.HttpListenerException (995): Der E/A-Vorgang wurde wegen eines Threadendes oder einer Anwendungsanforderung abgebrochen.
   at System.Net.HttpListener.GetContext()
   at CompuMaster.Web.TinyWebServerAdvanced.WebServer.<Run>b__20_0(Object o)
   at System.Threading.QueueUserWorkItemCallback.Execute()
   at System.Threading.ThreadPoolWorkQueue.Dispatch()
   at System.Threading.PortableThreadPool.WorkerThread.WorkerThreadStart()
```
* Npgsql
  * IST: v8.0
  * NEU: v9.0 -> ohne Support für .NET Framework 4.8 -> benötigt zuerst vollständiges Upgrade auf .NET (Core)
* NUnit
  * IST: v3.14
  * NEU: v4.x -> Refactoring aller Unit-Tests notwendig
* System.Data.SqlClient
  * IST: v4.8.6
  * NEU: v4.9.0 -> PlatformNotSupportedException: lost supported on netcoreapp3.1 platform