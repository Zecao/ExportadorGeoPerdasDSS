# ExportadorGeoPerdasDSS
This is a C# project which connects in GeoPerdasANEEL database (MS SQL Server 2014) and batch generates the OpenDSS files (.dss) for the distribution feeders in the exact format to be executed by the project https://github.com/Zecao/ExecutorOpenDssBr. 

The GeoPerdasANEEL is a database created from another database know as BDGD which Brazilian distribution utilities must annualy inform to the Brazilian regulatory agency - ANEEL. On these database the agency calculates the feeders energy losses.

Advantages of using this project: 
- faster .dss files creation, e.g. it takes 40 minutes to create 1700 feeders X 18 hours in the ETL tool.
- alternative of using GeoPerdas.EXE (the feeders .dss files - see FeederExample directory - are very similar to the ones created by the program GeoPerdas.exe, but they are customized to be executed by the C# project cited above. 

There are some improvements to do yet, like:
- It doesn't generates the "linecode" files and the load profiles files for the loads. As I use a fixed files for all feeders,  I've included these files in FeederExample directory.

The project also uses one external dll (EPPlus in lib directory) that allows Excel files to be read in the C#. 

Feel free to make contact if this project have some use for you or your company.
Ezequiel
