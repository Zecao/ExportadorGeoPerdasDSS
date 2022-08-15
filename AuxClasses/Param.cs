using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportadorGeoPerdasDSS
{
    class Param
    {
        public string _path; // main path
        public string _permRes; // persistent resources subdirectory
        public string _codBase; // ANEEL company number
        public string _pathAlim; // feeder subdirectory
        public string _conjAlim; // string to concatenates more than one feeder separated by ','
        public string _trEM; //trecho energy meter

        public Param(string path, string permRes, string codBase, string pathAlim, string conjAlim, string trEM)
        {
            _path = path;
            _permRes = permRes;
            _codBase = codBase;
            _pathAlim = pathAlim;
            _conjAlim = conjAlim;
            _trEM = trEM;
        }
    }
}
