# ExportadorGeoPerdasDSS
This is a C# data base project which connects in the database GeoPerdas(MS SQL Server 2014) and batch generates the OpenDSS files for the feeders, in the exact format to be executed by the project https://github.com/Zecao/ExecutorOpenDssBr, eliminating the step number 3, that is, the need of an ETL tool to create these .dss files. The main advantage of this project is .dss file creation is faster and takes less than half a hour (before it took 18hours in the ETL tool).

The "GeoPerdas" is a database where brazilian eletric distribution utilities must inform their feeders and assets for the regulatory agency ANEEL calculates the feeders energy losses. So, the feeders .dss files (see FeederExample directory) are very similar to the ones created by the  agency program (GeoPerdas.exe), but they are customized to be executed by the C# project cited above.

There are some improvements to do yet, like:
- In the moment, it works only with 13.8kV feeders;
- It generates only one month for the load files (you can choose the month); 
- It doesn't generates the "linecode" files and the "profiles" files for the loads (as I use a fixed files for all feeders, but Im including these in this version).

The project uses one external dll (EPPlus in lib directory) that allows Excel files to be read in the C#. 

Feel free to make contact if this project have some use for you or your company.
Ezequiel

