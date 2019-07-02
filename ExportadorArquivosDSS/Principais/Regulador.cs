using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class Regulador
    {
        // membros privados
        private static readonly string _reguladores = "Reguladores.dss";
        private StringBuilder _arqReguladorMT;

        private readonly string _codBase;
        private readonly string _pathAlim;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        
        public Regulador(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;

            _codBase = codBase;

        }

        private string GetNomeArqReguladorMT(string alim)
        {
            return _pathAlim + alim + _reguladores;
        }

        // new transformer.TRTR2332AN Phases = 1, windings = 2, buses = (BMT98402313.1.0, BMT165130397.1.0), conns = (LN, LN), kvs = (7.97 7.97), kvas = (1992, 1992), xhl = 0.75,%loadloss=0.125251004016064,%noloadloss=0.0268072289156626
        // new RegControl.RRTR2332AN transformer = TRTR2332AN, winding = 2, PTphase = 1, ptratio = 66.4, band = 3, vreg = 125
        public bool ConsultaStoredReguladorMT()
        {
            _arqReguladorMT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                // TODO incluir TnsLnh1_kV no select (e no banco em casa).
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodRegul,CodFasPrim,CodPonAcopl1,CodPonAcopl2,PotPass_kVA,[ReatHL_%],[Resis_%],PerdVz_W,TenRgl_pu,CodBnc,TipRegul from dbo.StoredReguladorMT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            //string tensaoFN = getTensaoFN(rs["TnsLnh1_kV"].ToString()); // TODO
                            string tensaoFN = AuxFunc.GetTensaoFN("13.8");
                            string tipoRegul = rs["TipRegul"].ToString();
                            string perVazioPer = CalcPerdVazio(rs);
                            string vRegVolts = CalcVReg(rs["TenRgl_pu"].ToString());
                            //string ptratio = getPTRatio(rs["TnsLnh1_kV"].ToString()); // TODO
                            string ptratio = GetPTRatio("13.8");

                            // banco de regulador
                            if (tipoRegul.Equals("4"))
                            {
                                string linha1 = "new transformer.RT" + rs["CodRegul"] + "-" + rs["CodBnc"].ToString() 
                                    + " Phases=1" 
                                    + ",windings=2" 
                                    + ",buses=[" + "BMT" + rs["CodPonAcopl1"] + "." + rs["CodBnc"].ToString() + ".0 " + "BMT" + rs["CodPonAcopl2"] + "." + rs["CodBnc"].ToString() + ".0]," //OBBS1
                                    + "conns=[LN LN]"
                                    + ",kvs=[" + tensaoFN + " " + tensaoFN + "]"
                                    + ",kvas=[" + rs["PotPass_kVA"].ToString() + " " + rs["PotPass_kVA"].ToString() + "]" 
                                    + ",xhl=" + rs["ReatHL_%"] 
                                    + ",%loadloss=" + rs["Resis_%"]
                                    + ",%noloadloss=" + perVazioPer + Environment.NewLine;

                                string linha2 = "new RegControl.RC" + rs["CodRegul"] + "-" + rs["CodBnc"].ToString() 
                                    + " transformer=RT" + rs["CodRegul"] + "-" + rs["CodBnc"].ToString() 
                                    + ",winding=2"
                                    + ",PTphase=1"
                                    + ",ptratio=" + ptratio 
                                    + ",band=2" 
                                    + ",vreg=" + vRegVolts + Environment.NewLine;

                                _arqReguladorMT.Append(linha1);
                                _arqReguladorMT.Append(linha2);

                            }
                            else
                            {
                                // TODO testar
                                string faseDSS = AuxFunc.GetFasesDSS(rs["CodFasPrim"].ToString());

                                string linha1 = "new transformer." + rs["CodRegul"] 
                                    + " Phases=1" + ",windings=2" 
                                    + ",buses=[" + "BMT" + rs["CodPonAcopl1"] + faseDSS + " " + "BMT" + rs["CodPonAcopl2"] + faseDSS + "]" //OBBS1
                                    + ",conns=[LN LN]" 
                                    + ",kvs=[" + tensaoFN + " " + tensaoFN + "]" 
                                    + ",kvas=[" + rs["PotPass_kVA"] + " "
                                    + rs[4] + "]" 
                                    + ",xhl=" + rs["ReatHL_%"] 
                                    + ",%loadloss=" + rs["Resis_%"] 
                                    + ",%noloadloss=" + perVazioPer + Environment.NewLine;

                                string linha2 = "new RegControl." + rs["CodRegul"] 
                                    + " transformer=" + rs["CodRegul"] 
                                    + ",winding=2" 
                                    + ",PTphase=1" 
                                    + ",ptratio=" + ptratio 
                                    + ",band=2"
                                    + ",vreg=" + vRegVolts + Environment.NewLine;

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
            double potNomKVA = double.Parse(rs["PotPass_kVA"].ToString());

            double perdaVazioPer = perVazioWatts / (potNomKVA * 10);

            return perdaVazioPer.ToString("0.####");
        }

        // get PT ratio de acordo com a tensao de linha
        private string GetPTRatio(string tensaoLinha)
        {
            string ret;

            switch (tensaoLinha)
            {
                case "34.5":
                    ret = "166.0";
                    break;
                case "22.0":
                    ret = "105.8";
                    break;
                case "13.8":
                    ret = "66.4";
                    break;
                default:
                    ret = "66.4";
                    break;
            }
            return ret;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ExecutorOpenDSS.ArqManip.GravaEmArquivo( _arqReguladorMT.ToString(), GetNomeArq());
        }

        private string GetNomeArq()
        {
            return _pathAlim + _alim + _reguladores;
        }
    }
}
