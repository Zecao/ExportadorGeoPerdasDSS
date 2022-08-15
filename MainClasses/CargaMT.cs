﻿using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportadorGeoPerdasDSS
{
    class CargaMT
    {
        // membros privados
        private static readonly string _cargaMT = "CargaMT_";
        private static int _iMes;
        private static string _ano;
        private Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoBT;
        private readonly List<List<int>> _numDiasFeriadoXMes;
        private readonly Dictionary<string, double> _somaCurvaCargaDiariaPU;
        private readonly Dictionary<string, List<string>> _curvasTipicasClientesMT;
        private readonly ModeloSDEE _SDEE;

        public CargaMT(string alim, SqlConnectionStringBuilder connBuilder, int iMes,
            string ano, List<List<int>> numDiasFeriadoXMes, Dictionary<string,double> somaCurvaCargaDiariaPU,
            ModeloSDEE sdee, Param par, Dictionary<string, List<string>> curvasCliMT = null)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
            _iMes = iMes;
            _ano = ano;
            _numDiasFeriadoXMes = numDiasFeriadoXMes;
            _somaCurvaCargaDiariaPU = somaCurvaCargaDiariaPU;
            _SDEE = sdee;

            if (_SDEE._utilizarCurvaDeCargaClienteMTIndividual) 
            {
                _curvasTipicasClientesMT = curvasCliMT;
            }
        }

        //modelo
        // new load.3009011004M1 bus1=BMT145105559.1.2.3.0,Phases=3,kv=13.8,kw=8.74290521528275,pf=0.92,Vminpu=0.93,Vmaxpu=1.5,model=2,daily=arqCurvaNormA4-3,status=variable
        // CodBase	CodConsMT	CodAlim	CodFas	CodPonAcopl	SemRedAssoc	TipCrvaCarga	EnerMedid01_MWh	EnerMedid02_MWh	EnerMedid03_MWh	EnerMedid04_MWh	EnerMedid05_MWh	EnerMedid06_MWh	EnerMedid07_MWh	EnerMedid08_MWh	EnerMedid09_MWh	EnerMedid10_MWh	EnerMedid11_MWh	EnerMedid12_MWh	Descr	CodSubAtrib	CodAlimAtrib	TnsLnh_kV	TnsFas_kV
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                // TODO add TnsLnh_kV 
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodConsMT,CodFas,CodPonAcopl,TipCrvaCarga,TnsLnh_kV,EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh," +
                        "EnerMedid05_MWh,EnerMedid06_MWh,EnerMedid07_MWh," +
                        "EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh from ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "dbo.StoredCargaMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "dbo.StoredCargaMT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            // Obtem o consumo de acordo com o mes
                            string consumoMes = AuxFunc.GetConsumoMesCorrente(rs,_iMes);

                            // se consumo nao eh vazio, transforma para double para verificar se zero
                            // OBS: optou-se por esta funcao visto que o banco pode retornar: "0","0.00000" e etc...
                            if (!consumoMes.Equals(""))
                            {
                                double dConsumoMes = double.Parse(consumoMes);

                                // skipa consumo = 0
                                if (dConsumoMes == 0)
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                continue;
                            }

                            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string tensaoFF = AuxFunc.GetTensaoFF(rs["TnsLnh_kV"].ToString()); 
                            
                            //
                            string demanda = AuxFunc.CalcDemanda(consumoMes, _iMes, _ano, rs["TipCrvaCarga"].ToString(), _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU);

                            string linha ="";

                            //
                            if (_SDEE._utilizarCurvaDeCargaClienteMTIndividual)
                            {

                                // se modelo de carga ANEEL
                                switch (_SDEE._modeloCarga)
                                {
                                    case "ANEEL":

                                        linha = CriaDSSCargaMTcomCurvaAneel(rs, demanda, fases, numFases, tensaoFF);

                                        break;

                                    case "PCONST":

                                        linha = CriaDSSCargaMTcomCurva(rs, demanda, fases, numFases, tensaoFF);

                                        break;
                                }
                            }
                            else 
                            {
                                // se modelo de carga ANEEL
                                switch (_SDEE._modeloCarga)
                                {
                                    case "ANEEL":

                                        linha = CriaDSSCargaMTAneel(rs, demanda, fases, numFases, tensaoFF);

                                        break;

                                    case "PCONST":

                                        linha = CriaDSSCargaPconst(rs, demanda, fases, numFases, tensaoFF);

                                        break;
                                }
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

        private string CriaDSSCargaMTcomCurvaAneel(SqlDataReader rs, string demanda, string fases, string numFases, string tensaoFF)
        {
            // divide demanda entre 2 cargas
            double demandaElse = double.Parse(demanda) / 2;

            string linha;

            string codCliMT = rs["CodConsMT"].ToString();

            // se cliente MT esta no dicionario de curvas de carga
            if (_curvasTipicasClientesMT.ContainsKey(codCliMT))
            {
                List<string> dadosCliMT = _curvasTipicasClientesMT[codCliMT];

                string fatorkdiario = dadosCliMT[3];

                // recalcula demanda base
                string demandaD = AuxFunc.CalcDemandaPorFatorKdiario(AuxFunc.GetConsumoMesCorrente(rs, _iMes), _iMes, _ano, fatorkdiario);

                // divide demanda entre 2 cargas
                double demanda2 = double.Parse(demandaD) / 2;

                // curva de carga
                linha = dadosCliMT[2] + Environment.NewLine;

                linha += "new load." + rs["CodConsMT"].ToString() + "M2"
                    + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                    + ",Phases=" + numFases
                    + ",kv=" + tensaoFF
                    + ",kW=" + demanda2.ToString("0.###")
                    + ",pf=" + dadosCliMT[1]
                    + ",Vminpu=0.93,Vmaxpu=1.5"
                    + ",model=2"
                    + ",daily=" + codCliMT
                    + ",status=variable" + Environment.NewLine;

                linha += "new load." + rs["CodConsMT"].ToString() + "M3"
                    + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                    + ",Phases=" + numFases
                    + ",kv=" + tensaoFF
                    + ",kW=" + demanda2.ToString("0.###")
                    + ",pf=" + dadosCliMT[1]
                    + ",Vminpu=0.93,Vmaxpu=1.5"
                    + ",model=3"
                    + ",daily=" + codCliMT
                    + ",status=variable" + Environment.NewLine;
            }
            else
            {
                linha = "new load." + rs["CodConsMT"].ToString() + "M2"
                    + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                    + ",Phases=" + numFases
                    + ",kv=" + tensaoFF
                    + ",kW=" + demandaElse.ToString()
                    + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                    + ",model=2"
                    + ",daily=" + rs["TipCrvaCarga"].ToString()
                    + ",status=variable" + Environment.NewLine;

                linha += "new load." + rs["CodConsMT"].ToString() + "M3"
                    + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                    + ",Phases=" + numFases
                    + ",kv=" + tensaoFF
                    + ",kW=" + demandaElse.ToString()
                    + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                    + ",model=3"
                    + ",daily=" + rs["TipCrvaCarga"].ToString()
                    + ",status=variable" + Environment.NewLine;
            }

            return linha;
        }

        private string CriaDSSCargaPconst(SqlDataReader rs, string demanda, string fases, string numFases, string tensaoFF)
        {
            string linha;
            double demandaD = double.Parse(demanda) * 2;

            linha = "new load." + rs["CodConsMT"].ToString() + "M1"
                + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                + ",Phases=" + numFases
                + ",kv=" + tensaoFF
                + ",kW=" + demandaD.ToString()
                + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                + ",model=1"
                + ",daily=" + rs["TipCrvaCarga"].ToString()
                + ",status=variable" + Environment.NewLine;

            return linha;
        }

        private string CriaDSSCargaMTAneel(SqlDataReader rs, string demanda, string fases, string numFases, string tensaoFF)
        {
            string linha;
            // carga model=2
            linha = "new load." + rs["CodConsMT"].ToString() + "M2"
                + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                + ",Phases=" + numFases
                + ",kv=" + tensaoFF
                + ",kW=" + demanda
                + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                + ",model=2"
                + ",daily=" + rs["TipCrvaCarga"].ToString()
                + ",status=variable" + Environment.NewLine;

            // carga model=3
            linha += "new load." + rs["CodConsMT"].ToString() + "M3"
                + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                + ",Phases=" + numFases
                + ",kv=" + tensaoFF
                + ",kW=" + demanda
                + ",pf=0.92,Vminpu=0.93,Vmaxpu=1.5"
                + ",model=3"
                + ",daily=" + rs["TipCrvaCarga"].ToString()
                + ",status=variable" + Environment.NewLine;
            
            return linha;
        }

        // OBS: cargas de MT criadas com PotCOnst
        private string CriaDSSCargaMTcomCurva(SqlDataReader rs, string demanda, string fases, string numFases, string tensaoFF)
        {
            string linha;
 
            string codCliMT = rs["CodConsMT"].ToString();

            // se cliente MT esta no dicionario de curvas de carga
            if (_curvasTipicasClientesMT.ContainsKey(codCliMT))
            {
                List<string> dadosCliMT = _curvasTipicasClientesMT[codCliMT];

                string fatorkdiario = dadosCliMT[3];

                // recalcula demanda base
                string demandaD = AuxFunc.CalcDemandaPorFatorKdiario(AuxFunc.GetConsumoMesCorrente(rs, _iMes), _iMes, _ano, fatorkdiario);

                // curva de carga
                linha = dadosCliMT[2] + Environment.NewLine;

                linha += "new load." + rs["CodConsMT"].ToString() + "M1"
                    + " bus1=" + "BMT" + rs["CodPonAcopl"] + fases //OBS1
                    + ",Phases=" + numFases
                    + ",kv=" + tensaoFF
                    + ",kW=" + demandaD.ToString()
                    + ",pf=" + dadosCliMT[1]
                    + ",Vminpu=0.93,Vmaxpu=1.5"
                    + ",model=1"
                    + ",daily=" + codCliMT
                    + ",status=variable" + Environment.NewLine;
            }
            else
            {
                    linha = CriaDSSCargaPconst(rs, demanda, fases, numFases, tensaoFF);
     
            }

            return linha;
        }

        public string GetNomeArq()
        {
            string strMes = AuxFunc.IntMes2strMes(_iMes);

            return _par._pathAlim + _alim + _cargaMT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
