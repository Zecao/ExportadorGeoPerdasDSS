# ExportadorGeoPerdasDSS
This is a C# project which connects in the database GeoPerdas(MS SQL Server 2014) and batch generates the OpenDSS files for the distribution feeders, in the exact format to be executed by the project https://github.com/Zecao/ExecutorOpenDssBr. The main advantage using this project is the faster dss files creation, as it takes less than half a hour (before it took 18hours in the ETL tool).

The "GeoPerdas" is a database where eletric distribution utilities from Brazil must annualy inform their feeders and assets for the regulatory agency, ANEEL, calculates the feeders energy losses. So, the feeders .dss files (see FeederExample directory) are very similar to the ones created by the  agency program (GeoPerdas.exe), but they are customized to be executed by the C# project cited above.

There are some improvements to do yet, like:
- It doesn't generates the "linecode" files and the load profiles files for the loads. As I use a fixed files for all feeders,  I've included these files in FeederExample directory.

The project also uses one external dll (EPPlus in lib directory) that allows Excel files to be read in the C#. 

Feel free to make contact if this project have some use for you or your company.
Ezequiel

