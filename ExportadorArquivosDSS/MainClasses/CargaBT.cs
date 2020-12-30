using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class CargaBT
    {
        // membros privados
        private static readonly string _cargaBT = "CargaBT_";
        private static int _iMes;
        private static string _ano;
        private Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoBT;
        private readonly List<List<int>> _numDiasFeriadoXMes;
        private readonly Dictionary<string, double> _somaCurvaCargaDiariaPU;
        private readonly ModeloSDEE _SDEE;

        public CargaBT(string alim, SqlConnectionStringBuilder connBuilder, int iMes, string ano,
            List<List<int>> numDiasFeriadoXMes, Dictionary<string, double> somaCurvaCargaDiariaPU, ModeloSDEE sdee,
            Param par)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
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
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText = "select TipTrafo,TenSecu_kV,CodConsBT,CodFas,CodPonAcopl,TipCrvaCarga,EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh,EnerMedid07_MWh," +
                            "EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh from " +
                            "dbo.StoredCargaBT as car inner join dbo.StoredTrafoMTMTMTBT as tr on tr.CodTrafo = car.CodTrafo " +
                            "where car.CodBase=@codbase and tr.CodBase=@codbase and car.CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = "select TipTrafo,TenSecu_kV,CodConsBT,CodFas,CodPonAcopl,TipCrvaCarga,EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh,EnerMedid07_MWh," +
                            "EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh from " +
                            "dbo.StoredCargaBT as car inner join dbo.StoredTrafoMTMTMTBT as tr on tr.CodTrafo = car.CodTrafo " +
                            "where car.CodBase=@codbase and tr.CodBase=@codbase and car.CodAlim=@CodAlim";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                        command.Parameters.AddWithValue("@CodAlim", _alim);
                    }

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

                            // se consumo nao eh vazio, transforma para double
                            // OBS: optou-se por esta funcao visto que o banco pode retornar: "0","0.00000" e etc...
                            if ( !consumoMes.Equals("") )
                            {
                                double dConsumoMes = double.Parse(consumoMes);

                                // skipa consumo = 0
                                if ( dConsumoMes == 0 )
                                {
                                    continue;
                                }
                            }
                            else 
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
                                       + ",status=variable";

                                    // carga model=3
                                    linha += "new load." + rs["CodConsBT"].ToString() + "M3"
                                        + " bus1=" + prefixoBarraBT + rs["CodPonAcopl"] + fases //OBS1
                                        + ",Phases=" + numFases
                                        + ",kv=" + Kv
                                        + ",kW=" + demanda
                                        + ",pf=0.92,Vminpu=0.92,Vmaxpu=1.5"
                                        + ",model=3"
                                        + ",daily=" + rs["TipCrvaCarga"].ToString()
                                        + ",status=variable";
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
                                       + ",status=variable"; 

                                    break;
                            }

                            // alterar numCust=0 p/ cargas do tipo IP (iluminacao publica)
                            if (rs["TipCrvaCarga"].ToString().Equals("IP"))
                            {
                                linha += ",NumCust=0" + Environment.NewLine;
                            }
                            else
                            {
                                linha += Environment.NewLine;
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
            string strMes = AuxFunc.IntMes2strMes(_iMes);

            return _par._pathAlim + _alim + _cargaBT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
