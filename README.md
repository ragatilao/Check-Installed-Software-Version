# Check-Installed-Software-Version
<a name="header1"></a>

Navigation
- [How to use this powershell script](#how-to-use)
- [License MIT](#license)
 
 
SQL Server Management Studio 17.0 came out in April 25, 2017. Since then, several version and iterations have been released by Microsoft. Now the latest version of SSMS is 17.8.1.

We have also built several database servers for the past year. Every time we build a server we also install the latest copy of the SQL Server Management Studio on the server.

After a year, our servers have different versions of SQL Server Management Studio on them. We did not do a good job of keeping and upgrading our SSMS on the servers.

Since we have several environments (Development, Test, Stage and Production) it was hard to keep up with the version.

This is not only designed for SQL Server Management Studio, it can also be used to check a specific software that are installed in multiple servers. 

## How to use this PowerShell script

* Run the PowerShell script. A window console will pop-up.
* Choose an option to enter a server name or use a server list file.
* Default value for software to be checked is SQL Server Management Studio, you can change this with other software name that you want to check.
* Click the Run Search button

[*Back to top*](#header1)

## License

[The Powershel script, Check Installed Software Version uses the MIT License.](LICENSE.md)

[*Back to top*](#header1)
