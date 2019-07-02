using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class Trafo
    {
        // membros privados
        private static string _trafos = "Transformadores.dss";
        private StringBuilder _arqTrafo;

        private string _codBase;
        private string _pathAlim;
        private string _alim;
        private static SqlConnectionStringBuilder _connBuilder;

        public Trafo(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;
            _codBase = codBase;
        }

        //new transformer.TRF1151404 Phases=1,Windings=3,Buses=[BMT162221270.3 BBT32992241.1.0 BBT32992241.0.2],Conns=[wye wye wye],kvs=[7.97 0.12 0.12],kvas=[15 15 15],Taps=[1.0 1 1],XscArray=[2.124,2.124,1.416],%loadloss=1.766666667 ,%noloadloss=0.433333333
        //new transformer.TRF1174742 Phases = 3, Windings = 2, Buses =[BMT98410048.1.2.3 BBT36975913.1.2.3.0], Conns =[delta wye], kvs =[13.80 0.220], kvas =[75 75], Taps =[1.0 1.0], XHL = 3.72,%loadloss=1.466666667 ,%noloadloss=0.393333333
        //CodBase	CodTrafo	CodBnc	CodAlim	CodPonAcopl1	CodPonAcopl2	PotNom_kVA	MRT	TipTrafo	CodFasPrim	CodFasSecu	CodFasTerc	TenSecu_kV	Tap_pu	Resis_%	ReatHL_%	ReatHT_%	ReatLT_%	PerdTtl_%	PerdVz_%	ClssTrafo	Propr	Descr	CodSubAtrib	CodAlimAtrib	Ordm	De	Para	TnsLnh1_kV	TnsLnh2_kV
        public bool ConsultaBanco()
        {
            _arqTrafo = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // TODO ADICIONAR TnsLnh1_kV	
                    command.CommandText = "select CodTrafo,CodPonAcopl1,CodPonAcopl2,PotNom_kVA,TipTrafo,CodFasPrim,CodFasSecu,CodFasTerc,TenSecu_kV,Tap_pu,[Resis_%],[ReatHL_%],[ReatHT_%]," +
                        "[ReatLT_%],[PerdVz_%] from dbo.StoredTrafoMTMTMTBT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string barraBT = rs["CodPonAcopl2"].ToString();
                            string pot = rs["PotNom_kVA"].ToString();
                            string tensaoFN = AuxFunc.GetTensaoFN("13.8");
                            string faseDSS = AuxFunc.GetFasesDSS(rs["CodFasPrim"].ToString());

                            // trafo mono
                            // OBS: verificado que no geoperdas2018, o trafo mono apresenta TipTrafo=1
                            if (rs["TipTrafo"].ToString().Equals("1")) 
                            {
                                string linha = "new transformer." + rs["CodTrafo"].ToString()
                                    + " Phases=1"
                                    + ",Windings=3"
                                    + ",Buses=[" + "BMT" + rs["CodPonAcopl1"].ToString() + faseDSS //OBS1
                                    + " " + "BBT" + barraBT + ".1.0 " + "BBT" + barraBT + ".0.2]"  //OBS1 //OBS: atenção para a polaridade .0.2
                                    + ",Conns=[wye wye wye]"             
                                    + ",kvs=[" + tensaoFN + " " + "0.12 0.12]"
                                    + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                                    + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]"
                                    + ",XHL=" + rs["ReatHL_%"].ToString()
                                    + ",XHT=" + rs["ReatHT_%"].ToString()
                                    + ",XLT=" + rs["ReatLT_%"].ToString() 
                                    + ",%loadloss=" + rs["Resis_%"].ToString()
                                    + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

                                _arqTrafo.Append(linha);
                            }
                            else
                            {
                                string linha = "new transformer." + rs["CodTrafo"].ToString()
                                     + " Phases=3"
                                     + ",Windings=2"
                                     + ",Buses=[" + "BMT" +rs["CodPonAcopl1"].ToString() + faseDSS //OBS1
                                     + " " + "BBT" + barraBT + ".1.2.3.0]" //OBS1
                                     + ",Conns=[delta wye]"
                                     + ",kvs=[" + "13.8" + " " + "0.22]" // TODO implementar funcao tensao
                                     + ",kvas=[" + pot + " " + pot + "]"
                                     + ",Taps=[1," + rs["Tap_pu"].ToString() + "]"
                                     + ",XHL=" + rs["ReatHL_%"].ToString()
                                     + ",%loadloss=" + rs["Resis_%"].ToString()
                                     + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

                                _arqTrafo.Append(linha);
                            }
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        private string GetNomeArq()
        {
            return _pathAlim + _alim + _trafos;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqTrafo.ToString(), GetNomeArq());
        }
    }
}
