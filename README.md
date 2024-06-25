# ExportadorGeoPerdasDSS
This is a C# project that connects in GeoPerdasANEEL database (MS SQL Server) and batch generates the OpenDSS files (.dss) for the distribution feeders in the exact format to be executed by the project https://github.com/Zecao/ExecutorOpenDssBr. The GeoPerdasANEEL is a database created from another database known as BDGD which the Brazilian distribution utilities annually must inform all the main assets to the Brazilian regulatory agency, ANEEL. So, the agency calculates the feeder's energy losses using OpenDSS and the files created from this database. The feeders .dss files - see FeederExample directory - are very similar to the ones created by the ANEEL software GeoPerdas.exe, but they are customized to be executed by the C# project cited above.

Advantages of using this project: 
- faster .dss file creation, e.g. it takes 40 minutes to create 1700 feeders X 18 hours using a commercial ETL tool.
- alternative of using the ANEEL software GeoPerdas.EXE. 
- addition of some customizations that do not exist in ANEEL .dss files, such as capacitor files, alternative to the ANEEL load model, modeling of all MV and LV Generators and PVSystems, reclosers and fuses. Of course, this information must also be available in the GeoPerdas database, through extensions of the original database.

The project also uses one external dll (EPPlus in lib directory) that allows Excel files to be read in the C#. 

### Updates in 01/08/2024:
- The current release now generates the LineCode .dss files. Please, see the function.
- Furthermore, in addition to Cemig GeoPerdas, I tested the project in 2 more BDGD/GeoPerdas from other Brazilian utilities (Equatorial and Neoenergia). It was necessary to make some adjustments for specific characteristics of these utilities (e.g. transformer secondary voltage level), but everything is in the code.

There are some improvements to make, like:
- Addition of the SQL Scripts to extend the original GeoPerdas database (e.g. recolosers and fuse modelling and capacitor table).  
- Generate the load profile files. As the LoadProfiles are the same for all feeders, it was easier to create them directly from a SQL query and use Excel to adjustments (e.g. to change the 96 points data to 24 points, etc.). I've included these files in the FeederExample directory.

Feel free to contact me if this project has some use for you or your company.
Ezequiel
