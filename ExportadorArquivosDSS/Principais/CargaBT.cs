using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class CargaBT
    {
        // membros privados
        private static string _cargaBT = "CargaBT_";
        private static int _iMes;
        private static string _ano;
        private readonly string _codBase;
        private readonly string _pathAlim;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoBT;
        private readonly List<List<int>> _numDiasFeriadoXMes;
        private readonly Dictionary<string, double> _somaCurvaCargaDiariaPU;
        private readonly ModeloSDEE _SDEE;

        public CargaBT(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase, int iMes, string ano, List<List<int>> numDiasFeriadoXMes, Dictionary<string, double> somaCurvaCargaDiariaPU, ModeloSDEE sdee)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;
            _codBase = codBase;
            _iMes = iMes;
            _ano = ano;
            _numDiasFeriadoXMes = numDiasFeriadoXMes;
            _somaCurvaCargaDiariaPU = somaCurvaCargaDiariaPU;
            _SDEE = sdee;
        }

        //modelo
        //new load.3001215463M1 bus1=R9772.1.3.0,Phases=2,kv=0.22,kw=1.29794758726823,pf=0.92,Vminpu=0.92,Vmaxpu=1.5,model=2,daily=arqCurvaNormRES4_11,status=variable
        //new load.3001215463M2 bus1=R9772.1.3.0,Phases=2,kv=0.22,kw=1.29794758726823,pf=0.92,Vminpu=0.92,Vmaxpu=1.5,model=3,daily=arqCurvaNormRES4_11,status=variable
        // CodBase	CodConsBT	CodAlim	CodTrafo	CodRmlBT	CodFas	CodPonAcopl	SemRedAssoc	TipMedi	TipCrvaCarga	EnerMedid01_MWh	EnerMedid02_MWh	EnerMedid03_MWh	EnerMedid04_MWh	EnerMedid05_MWh	EnerMedid06_MWh	EnerMedid07_MWh	EnerMedid08_MWh	EnerMedid09_MWh	EnerMedid10_MWh	EnerMedid11_MWh	EnerMedid12_MWh	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	TnsLnh_kV	TnsFas_kV
        public bool ConsultaBanco()
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select TipTrafo,TenSecu_kV,CodConsBT,CodFas,CodPonAcopl,TipCrvaCarga,EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh,EnerMedid07_MWh," +
                        "EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh from " +
                        "dbo.StoredCargaBT as car inner join dbo.StoredTrafoMTMTMTBT as tr on tr.CodTrafo = car.CodTrafo " +
                        "where car.CodBase=@codbase and car.CodAlim=@CodAlim";
                    command.Parameters.AddWithValue("@codbase", _codBase);
                    command.Parameters.AddWithValue("@CodAlim", _alim);

                    using (var rs = command.ExecuteReader())
                    {
                        // verifica ocorrencia de elemento no banco
                        if (!rs.HasRows)
                        {
                            return false;
                        }

                        while (rs.Read())
                        {
                            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string prefixoBarraBT = GetPrefixoBarraBT(rs["CodConsBT"].ToString());

                            // obtem tensao base de acordo com tipo da carga (mono, bi ou tri) e o nivel de tensao do tipo do trafo
                            string Kv = GetTensaoBase(numFases, rs["TipTrafo"].ToString());

                            //obtem o consumo de acordo com o mes 
                            string consumoMes = AuxFunc.GetConsumoMesCorrente(rs,_iMes);

                            // skipa consumo = 0
                            if ( consumoMes.Equals("0") || consumoMes.Equals("") )
                            {
                                continue;
                            }

                            string demanda = AuxFunc.CalcDemanda(consumoMes, _iMes, _ano, rs["TipCrvaCarga"].ToString(), _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU);

                            string linha = "";
                            // se modelo de carga ANEEL
                            switch (_SDEE._modeloCarga)
                            {
                                case "ANEEL":

                                    // carga model=2
                                    linha = "new load." + rs["CodConsBT"].ToString() + "M2"
                                       + " bus1=" + prefixoBarraBT + rs["CodPonAcopl"] + fases //OBS1
                                       + ",Phases=" + numFases
                                       + ",kv=" + Kv
                                       + ",kW=" + demanda
                                       + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                                       + ",model=2"
                                       + ",daily=" + rs["TipCrvaCarga"].ToString()
                                       + ",status=variable" + Environment.NewLine;

                                    // carga model=3
                                    linha += "new load." + rs["CodConsBT"].ToString() + "M3"
                                        + " bus1=" + prefixoBarraBT + rs["CodPonAcopl"] + fases //OBS1
                                        + ",Phases=" + numFases
                                        + ",kv=" + Kv
                                        + ",kW=" + demanda
                                        + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                                        + ",model=3"
                                        + ",daily=" + rs["TipCrvaCarga"].ToString()
                                        + ",status=variable" + Environment.NewLine;
                                    break;

                                // modelo P constante
                                case "PCONST":

                                    double demandaD = double.Parse(demanda) * 2;

                                    linha = "new load." + rs["CodConsBT"].ToString() + "M1"
                                       + " bus1=" + prefixoBarraBT + rs["CodPonAcopl"] + fases //OBS1
                                       + ",Phases=" + numFases
                                       + ",kv=" + Kv
                                       + ",kW=" + demandaD.ToString()
                                       + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                                       + ",model=1"
                                       + ",daily=" + rs["TipCrvaCarga"].ToString()
                                       + ",status=variable" + Environment.NewLine;

                                    break;
                            }

                            _arqSegmentoBT.Append(linha);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        // A tensao base depende do numero de fases e tambem do trafo.
        private string GetTensaoBase(string numFases, string tipoTrafo)
        {
            string retFases;
            switch (numFases)
            {
                case "2":
                    
                    // se trafo monofasico com tap central
                    if (tipoTrafo.Equals("1")) // bifasico do center tap
                        retFases = "0.24";
                    else
                        retFases = "0.22";
                    break;

                case "1":

                    // se trafo monofasico com tap central
                    if (tipoTrafo.Equals("1")) // monofasico do center tap
                        retFases = "0.12";
                    else
                        retFases = "0.127";
                    break;

                default: // se trifasico, tensao base BT sempre sera 0.22 
                    retFases = "0.22";
                    break;
            }
            return retFases;
        }

        // Distingue carga BT de IP e retorna o prefixo da Barra de BT correto. 
        private string GetPrefixoBarraBT(string v)
        {
            if (v.Contains("IP"))
                return "BBT";
            else
                return "RML";
        }

        public string GetNomeArq()
        {
            string strMes = AuxFunc.intMes2strMes(_iMes);

            return _pathAlim + _alim + _cargaBT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
