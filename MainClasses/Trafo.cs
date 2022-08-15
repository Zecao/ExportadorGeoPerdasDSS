using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportadorGeoPerdasDSS
{
    class Trafo
    {
        // membros privados
        private static readonly string _trafos = "Transformadores.dss";
        private StringBuilder _arqTrafo;
        private Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private ModeloSDEE _SDEE;

        public Trafo(string alim, SqlConnectionStringBuilder connBuilder, Param par, ModeloSDEE SDEE)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
            _SDEE = SDEE;
        }

        //new transformer.TRF1151404 Phases=1,Windings=3,Buses=[BMT162221270.3 BBT32992241.1.0 BBT32992241.0.2],Conns=[wye wye wye],kvs=[7.97 0.12 0.12],kvas=[15 15 15],Taps=[1.0 1 1],XscArray=[2.124,2.124,1.416],%loadloss=1.766666667 ,%noloadloss=0.433333333
        //new transformer.TRF1174742 Phases = 3, Windings = 2, Buses =[BMT98410048.1.2.3 BBT36975913.1.2.3.0], Conns =[delta wye], kvs =[13.80 0.220], kvas =[75 75], Taps =[1.0 1.0], XHL = 3.72,%loadloss=1.466666667 ,%noloadloss=0.393333333
        //CodBase	CodTrafo	CodBnc	CodAlim	CodPonAcopl1	CodPonAcopl2	PotNom_kVA	MRT	TipTrafo	CodFasPrim	CodFasSecu	CodFasTerc	TenSecu_kV	Tap_pu	Resis_%	ReatHL_%	ReatHT_%	ReatLT_%	PerdTtl_%	PerdVz_%	ClssTrafo	Propr	Descr	CodSubAtrib	CodAlimAtrib	Ordm	De	Para	TnsLnh1_kV	TnsLnh2_kV
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqTrafo = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select Propr,CodTrafo,CodPonAcopl1,CodPonAcopl2,PotNom_kVA,TipTrafo,CodFasPrim,CodFasSecu,CodFasTerc,TenSecu_kV,"
                        + "Tap_pu,[Resis_%],[ReatHL_%],[ReatHT_%],[ReatLT_%],[PerdVz_%],TnsLnh1_kV,CodBnc,Descr"
                        + " from dbo.StoredTrafoMTMTMTBT ";

                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText += "where CodBase=@codbase and CodAlim=@CodAlim";
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
                            //skipa trafos de terceiros
                            if (rs["Propr"].ToString().Equals("TC"))
                            {
                                // Console.Write("trafo de terceiros: " + rs["CodTrafo"].ToString() + Environment.NewLine );
                                continue;
                            }

                            string barraBT = rs["CodPonAcopl2"].ToString();
                            string pot = rs["PotNom_kVA"].ToString();

                            // double pot2 =Double.TryParse(rs["PotNom_kVA"];
                            // pot = string.Format("{0:.##}}", pot2);

                            string tensaoFF = rs["TnsLnh1_kV"].ToString(); 
                            
                            string tensaoFN = AuxFunc.GetTensaoFN(tensaoFF);
                            string faseDSS = AuxFunc.GetFasesDSS(rs["CodFasPrim"].ToString());

                            string CodTrafo = rs["CodTrafo"].ToString();

                            string linha;

                            // OBS: verificado que no geoperdas2018, o trafo mono apresenta TipTrafo=1
                            switch (rs["TipTrafo"].ToString())
                            {
                                //monofasico 3 fios
                                case "1":
                                    linha = CriaStringTrafoMonofasico(rs, faseDSS, barraBT, tensaoFN, pot);
                                    break;
                                //monofasico 3 fios
                                case "2":
                                    linha = CriaStringTrafoMonofasico(rs, faseDSS, barraBT, tensaoFN, pot);
                                    break;
                                //posto transformador
                                case "5":
                                    linha = CriaStringPostoTransformador(rs, faseDSS, barraBT, tensaoFF, pot);
                                    break;
                                default:
                                    linha = CriaStringTrafoTrifasico(rs, faseDSS, barraBT, tensaoFF, pot);
                                    break;
                            }
                            _arqTrafo.Append(linha);
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }
            return true;
        }

        //
        private string TensaoLinha2TensaoFase(string tensaoFF)
        {
            /*
            //verifica se eh vazia 
            if ( tensaoFF.Equals("") && tipoBanco.Equals("5") )
            {
                return "34.50";
            }
            if ( tensaoFF.Equals("") && tipoBanco.Equals("4") )
            {
                return "13.80";
            }*/

            //necessario transformar para double
            double tensaoFFd = double.Parse(tensaoFF);

            // Se tensao fasePrimDSS, o posto eh abaixador 
            if (tensaoFFd.Equals(34.5))
            {
                return "19.92";
            }
            if (tensaoFFd.Equals(13.8))
            {
                return "7.97";
            }
            return "7.97";
        }


        /* Exemplo de declaracao PT
! !elevador
new transformer.873922PT_1 Phases=1,Windings=2,Buses=[BMT174122367.1.2 BMT174122366.1.0],Conns=[delta wye],kvs=[13.8  19.92]
! !abaixador 
new transformer.873969PT_1 Phases=1,Windings=2,Buses=[BMT174122333.1.2 BMT174122332.1.0],Conns=[delta wye],kvs=[34.5 7.97]
        */

        private string CriaStringPostoTransformador(SqlDataReader rs, string faseDSS, string barraBT, string fasePrimDSS, string pot)
        {
            string faseSecuDSS = AuxFunc.GetFasesDSS(rs["CodFasSecu"].ToString());

            // OBS: de acordo com o padrao do GeoPerdas, o CodFasSec Como o PT eh delta/estrela, a tensao secundaria de linha deve ser transformada em tensao de fase.
            string tensaoSec = TensaoLinha2TensaoFase(rs["TenSecu_kV"].ToString()); 

            // 
            string linha = "new transformer." + rs["CodTrafo"].ToString() + "_" + rs["CodBnc"].ToString()
                 + " Phases=1"
                 + ",Windings=2"
                 + ",Buses=[" + "BMT" + rs["CodPonAcopl1"].ToString() + faseDSS 
                 + " " + "BMT" + barraBT + faseSecuDSS + "]" 
                 + ",Conns=[delta wye]"
                 + ",kvs=[" + fasePrimDSS + " " + tensaoSec  + "]" 
                 + ",kvas=[" + pot + " " + pot + "]"
                 + ",Taps=[1," + rs["Tap_pu"].ToString() + "]";

            // se modo reatancia
            if (_SDEE._reatanciaTrafos)
            {
                linha += ",XHL=" + rs["ReatHL_%"].ToString();
            }

            linha += ",%loadloss=" + rs["Resis_%"].ToString()
                 + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

            return linha;
        }

        // cria string trafo trifasico
        private string CriaStringTrafoMonofasico(SqlDataReader rs, string faseDSS, string barraBT, string tensaoFN, string pot)
        {
            string linha = "new transformer." + rs["CodTrafo"].ToString()
                + " Phases=1"
                + ",Windings=3"
                + ",Buses=[" + "BMT" + rs["CodPonAcopl1"].ToString() + faseDSS //OBS1
                + " " + "BBT" + barraBT + ".1.0 " + "BBT" + barraBT + ".0.2]"  //OBS1 //OBS: atenção para a polaridade .0.2
                + ",Conns=[wye wye wye]"
                + ",kvs=[" + tensaoFN + " " + "0.12 0.12]"
                + ",kvas=[" + pot + " " + pot + " " + pot + "]"
                + ",Taps=[1," + rs["Tap_pu"].ToString() + "," + rs["Tap_pu"].ToString() + "]";

            // se modo reatancia
            if (_SDEE._reatanciaTrafos)
            {
                linha += ",XHL=" + rs["ReatHL_%"].ToString()
                + ",XHT=" + rs["ReatHT_%"].ToString()
                + ",XLT=" + rs["ReatLT_%"].ToString();
            }

            linha += ",%loadloss=" + rs["Resis_%"].ToString()
                + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

            return linha;
        }
        
        // cria string trafo trifasico
        private string CriaStringTrafoTrifasico(SqlDataReader rs, string faseDSS, string barraBT, string tensaoFF, string pot)
        {
            // TODO OBS: comentado pois alterava a tensao de linha correta do PT 34,5 kV 
            //tensaoFF = TrataTensaoLinhaANEEL(tensaoFF,"4");

            //
            string linha = "new transformer." + rs["CodTrafo"].ToString()
                 + " Phases=3"
                 + ",Windings=2"
                 + ",Buses=[" + "BMT" + rs["CodPonAcopl1"].ToString() + faseDSS //OBS1
                 + " " + "BBT" + barraBT + ".1.2.3.0]" //OBS1
                 + ",Conns=[delta wye]"
                 + ",kvs=[" + tensaoFF + " " + "0.22]"
                 + ",kvas=[" + pot + " " + pot + "]"
                 + ",Taps=[1," + rs["Tap_pu"].ToString() + "]";

            // se modo reatancia
            if (_SDEE._reatanciaTrafos)
            {
                linha += ",XHL=" + rs["ReatHL_%"].ToString();
            }

            linha += ",%loadloss=" + rs["Resis_%"].ToString()
                 + ",%noloadloss=" + rs["PerdVz_%"].ToString() + Environment.NewLine;

            return linha;
        }

        private string GetNomeArq()
        {
            return _par._pathAlim + _alim + _trafos;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqTrafo.ToString(), GetNomeArq());
        }
    }
}
