
namespace ExportadorGeoPerdasDSS
{
    class Param
    {
        public string _schema; // database
        public string _path; // main path
        public string _permRes; // persistent resources subdirectory
        public string _codBase; // ANEEL company number
        public string _pathAlim; // feeder subdirectory
        public string _conjAlim; // string to concatenates more than one feeder separated by ','
        public string _trEM; //trecho energy meter
        public bool _modelo4condutores; // 4 model

        public Param(string path, string permRes, string codBase, string pathAlim, string conjAlim, string trEM, bool modelo4condutores, string schema)
        {
            _schema = schema;
            _path = path;
            _permRes = permRes;
            _codBase = codBase;
            _pathAlim = pathAlim;
            _conjAlim = conjAlim;
            _trEM = trEM;
            _modelo4condutores = modelo4condutores;
        }
    }
}
