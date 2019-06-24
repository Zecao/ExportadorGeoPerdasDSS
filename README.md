# ExportadorGeoPerdasDSS
This is a C# data base project which connects in the database GeoPerdas(MS SQL Server 2014) and batch generates the OpenDSS files from real medium voltage feeders, in the exact format to be executed by the project https://github.com/Zecao/ExecutorOpenDssBr, eliminating the step number 3, that is, the need of an ETL tool to create these .dss files. 

The main advantage is as the file creation from more than one thousand feeders tooks a lot of time, now I can create these files in  less than half a hour.   

The feeder *.dss files (see FeederExample directory) are very similar to the ones created by the brazilian regulatory agency ANEEL
program (GeoPerdas.exe), but they are customized to be executed by the project cited above.
