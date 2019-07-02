using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class CargaMT
    {
        // membros privados
        private static string _cargaMT = "CargaMT_";
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

        public CargaMT(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase, int iMes, string ano, List<List<int>> numDiasFeriadoXMes, Dictionary<string,double> somaCurvaCargaDiariaPU, ModeloSDEE sdee)
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
        // new load.3009011004M1 bus1=BMT145105559.1.2.3.0,Phases=3,kv=13.8,kw=8.74290521528275,pf=0.92,Vminpu=0.93,Vmaxpu=1.5,model=2,daily=arqCurvaNormA4-3,status=variable
        // CodBase	CodConsMT	CodAlim	CodFas	CodPonAcopl	SemRedAssoc	TipCrvaCarga	EnerMedid01_MWh	EnerMedid02_MWh	EnerMedid03_MWh	EnerMedid04_MWh	EnerMedid05_MWh	EnerMedid06_MWh	EnerMedid07_MWh	EnerMedid08_MWh	EnerMedid09_MWh	EnerMedid10_MWh	EnerMedid11_MWh	EnerMedid12_MWh	Descr	CodSubAtrib	CodAlimAtrib	TnsLnh_kV	TnsFas_kV
        public bool ConsultaBanco()
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                // TODO add TnsLnh_kV 
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodConsMT,CodFas,CodPonAcopl,TipCrvaCarga,EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh,EnerMedid07_MWh," +
                        "EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh from " +
                        "dbo.StoredCargaMT where CodBase=@codbase and CodAlim=@CodAlim";
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

                            // Obtem o consumo de acordo com o mes
                            string consumoMes = AuxFunc.GetConsumoMesCorrente(rs,_iMes);

                            // skipa consumo = 0
                            if (consumoMes.Equals("0"))
                            {
                                continue;
                            }

                            string demanda = AuxFunc.CalcDemanda(consumoMes, _iMes, _ano, rs["TipCrvaCarga"].ToString(), _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU);

                            string linha ="";
                            // se modelo de carga ANEEL
                            switch (_SDEE._modeloCarga)
                            {
                                case "ANEEL":

                                    // carga model=2
                                    linha = "new load." + rs["CodConsMT"].ToString() + "M2"
                                        + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                                        + ",Phases=" + numFases
                                        + ",kv=13.8"
                                        + ",kW=" + demanda
                                        + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                                        + ",model=2"
                                        + ",daily=" + rs["TipCrvaCarga"].ToString()
                                        + ",status=variable" + Environment.NewLine;

                                    // carga model=3
                                    linha += "new load." + rs["CodConsMT"].ToString() + "M3"
                                        + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                                        + ",Phases=" + numFases
                                        + ",kv=13.8"
                                        + ",kW=" + demanda
                                        + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                                        + ",model=3"
                                        + ",daily=" + rs["TipCrvaCarga"].ToString()
                                        + ",status=variable" + Environment.NewLine;
                                    break;

                                case "PCONST":

                                    double demandaD = double.Parse(demanda) * 2;

                                    linha = "new load." + rs["CodConsMT"].ToString() + "M1"
                                        + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                                        + ",Phases=" + numFases
                                        + ",kv=13.8"
                                        + ",kW=" + demandaD.ToString()
                                        + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
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



        public string GetNomeArq()
        {
            string strMes = AuxFunc.intMes2strMes(_iMes);

            return _pathAlim + _alim + _cargaMT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
