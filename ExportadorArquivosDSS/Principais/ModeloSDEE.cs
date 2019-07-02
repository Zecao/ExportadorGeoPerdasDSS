using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    // opcoes de modelos SDEE 
    class ModeloSDEE
    {
        public bool _usarCondutoresSeqZero = false;
        public bool _utilizarCurvaDeCargaClienteMTIndividual = true;
        public bool _incluirCapacitoresMT = false;
        public string _modeloCarga = "ANEEL"; // Modelos disponiveis: "ANEEL" "PCONST"
        //public bool _incluirReatanciaDispersaoTrafos = true;

        public ModeloSDEE(bool usarCondutoresSeqZero, bool utilizarCurvaDeCargaClienteMTIndividual, bool incluirCapacitoresMT, string modeloCarga)
        {
            _usarCondutoresSeqZero = usarCondutoresSeqZero;
            _utilizarCurvaDeCargaClienteMTIndividual = utilizarCurvaDeCargaClienteMTIndividual;
            _incluirCapacitoresMT = incluirCapacitoresMT;
            _modeloCarga = modeloCarga;
            //_incluirReatanciaDispersaoTrafos = p3; //TODO
        }
    }
}
