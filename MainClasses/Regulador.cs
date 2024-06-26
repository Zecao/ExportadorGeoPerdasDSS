﻿using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class Regulador
    {
        // membros privados
        private static readonly string _reguladores = "Reguladores.dss";
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqReguladorMT;
        private Param _par;
        private readonly string _alim;

        public Regulador(SqlConnectionStringBuilder connBuilder, Param par)
        {
            _par = par;
            _alim = par._alim;
            _connBuilder = connBuilder;
        }

        private string GetNomeArqReguladorMT(string alim)
        {
            return _par._pathAlim + alim + _reguladores;
        }

        // new transformer.TRTR2332AN Phases = 1, windings = 2, buses = (BMT98402313.1.0, BMT165130397.1.0), conns = (LN, LN), kvs = (7.97 7.97), kvas = (1992, 1992), xhl = 0.75,%loadloss=0.125251004016064,%noloadloss=0.0268072289156626
        // new RegControl.RRTR2332AN transformer = TRTR2332AN, winding = 2, PTphase = 1, ptratio = 66.4, band = 3, vreg = 125
        public bool ConsultaStoredReguladorMT(bool _modoReconf)
        {
            _arqReguladorMT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText = "select CodRegulMT,TnsLnh1_kV,CodFasPrim,CodPonAcopl1,CodPonAcopl2,PotNom_kVA,[ReatHL_%],"
                            + "[Resis_%],PerdVz_W,TenRgl_pu,CodBnc,TipRegul,Descr from " + _par._DBschema + "StoredReguladorMT "
                            + "where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = "select CodRegulMT,TnsLnh1_kV,CodFasPrim,CodPonAcopl1,CodPonAcopl2,PotNom_kVA,[ReatHL_%],[Resis_%],PerdVz_W,TenRgl_pu,CodBnc,TipRegul,Descr from " +
                            _par._DBschema + "StoredReguladorMT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            // OBS: necessario converter a TnsLnh1_kV p/ fase-neutro, visto as mesmas ja virem adequadas p/ o caso do BRT em estrela  
                            string tensaoFN = AuxFunc.GetTensaoFN(rs["TnsLnh1_kV"].ToString());
                            string ptratio = GetPTRatio(tensaoFN);
                            string tipoRegul = rs["TipRegul"].ToString();
                            string perVazioPer = CalcPerdVazio(rs);
                            string vRegVolts = CalcVReg(rs["TenRgl_pu"].ToString());
                            string numEq = " ! " + rs["Descr"].ToString();

                            // banco de regulador
                            if (tipoRegul.Equals("4"))
                            {
                                string linha1 = "new transformer.RT" + rs["CodRegulMT"] + "-" + rs["CodBnc"].ToString()
                                    + " Phases=1"
                                    + ",windings=2"
                                    + ",buses=[" + "BMT" + rs["CodPonAcopl1"] + "." + rs["CodBnc"].ToString() + ".0 " + "BMT" + rs["CodPonAcopl2"] + "." + rs["CodBnc"].ToString() + ".0]," //OBBS1
                                    + "conns=[LN LN]"
                                    + ",kvs=[" + tensaoFN + " " + tensaoFN + "]"
                                    + ",kvas=[" + rs["PotNom_kVA"].ToString() + " " + rs["PotNom_kVA"].ToString() + "]"
                                    + ",xhl=" + rs["ReatHL_%"]
                                    + ",%loadloss=" + rs["Resis_%"]
                                    + ",%noloadloss=" + perVazioPer + Environment.NewLine;

                                string linha2 = "new RegControl.RC" + rs["CodRegulMT"] + "-" + rs["CodBnc"].ToString()
                                    + " transformer=RT" + rs["CodRegulMT"] + "-" + rs["CodBnc"].ToString()
                                    + ",winding=2"
                                    + ",PTphase=1"
                                    + ",ptratio=" + ptratio
                                    + ",band=3"
                                    + ",vreg=" + vRegVolts
                                    + ",revNeutral=Yes " // Does not regulate in reverse direction
                                    + numEq + Environment.NewLine;
                                /*
                                + ",reversible=Yes,revband=3"
                                + ",revvreg=" + "120"//vRegVolts //TODO
                                //+ ",revThreshold=10"*/

                                _arqReguladorMT.Append(linha1);
                                _arqReguladorMT.Append(linha2);

                            }
                            else
                            {
                                // TODO testar
                                string faseDSS = AuxFunc.GetFasesDSS(rs["CodFasPrim"].ToString());

                                string linha1 = "new transformer." + rs["CodRegulMT"]
                                    + " Phases=1" + ",windings=2"
                                    + ",buses=[" + "BMT" + rs["CodPonAcopl1"] + faseDSS + " " + "BMT" + rs["CodPonAcopl2"] + faseDSS + "]" //OBBS1
                                    + ",conns=[LN LN]"
                                    + ",kvs=[" + tensaoFN + " " + tensaoFN + "]"
                                    + ",kvas=[" + rs["PotNom_kVA"] + " " + rs["PotNom_kVA"] + "]"
                                    + ",xhl=" + rs["ReatHL_%"]
                                    + ",%loadloss=" + rs["Resis_%"]
                                    + ",%noloadloss=" + perVazioPer + Environment.NewLine;

                                string linha2 = "new RegControl." + rs["CodRegulMT"]
                                    + " transformer=" + rs["CodRegulMT"]
                                    + ",winding=2"
                                    + ",PTphase=1"
                                    + ",ptratio=" + ptratio
                                    + ",band=3"
                                    + ",vreg=" + vRegVolts
                                    + ",reversible=Yes,revband=3"
                                    + ",revvreg=" + vRegVolts
                                    //+ ",revThreshold=10" 
                                    + Environment.NewLine;

                                _arqReguladorMT.Append(linha1);
                                _arqReguladorMT.Append(linha2);
                            }
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        // calcula VReg em volts
        private string CalcVReg(string tensaoPUstr)
        {
            double tensaoPU = double.Parse(tensaoPUstr);
            double tensaoVolts = 120 * tensaoPU;

            return tensaoVolts.ToString("0.###");
        }

        // calcula perda vazio percentual
        private string CalcPerdVazio(SqlDataReader rs)
        {
            double perVazioWatts = double.Parse(rs["PerdVz_W"].ToString());
            double potNomKVA = double.Parse(rs["PotNom_kVA"].ToString());

            double perdaVazioPer = perVazioWatts / (potNomKVA * 10);

            return perdaVazioPer.ToString("0.####");
        }

        // get PT ratio de acordo com a tensao de linha
        private string GetPTRatio(string tensaoLinha)
        {
            //relacao TP default para RT de 7.97kV
            string ret = "66.4"; ;

            // retorno funcao para RT nao alcancado pela seq. eletrica.
            if (tensaoLinha.Equals(""))
            {
                return ret;
            }

            double tensaoFFd = double.Parse(tensaoLinha);

            // verifica se o nivel de tensaaFF eh 34.5kV ou 22.0kV
            if (tensaoFFd.Equals(19.92))
            {
                ret = "166.0";
            }
            else if (tensaoFFd.Equals(12.70))
            {
                ret = "105.8";
            }
            return ret;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqReguladorMT.ToString(), GetNomeArq());
        }

        private string GetNomeArq()
        {
            return _par._pathAlim + _alim + _reguladores;
        }
    }
}
