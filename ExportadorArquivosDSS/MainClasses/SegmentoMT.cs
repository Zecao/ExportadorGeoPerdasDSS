using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class SegmentoMT
    {
        // membros privados
        private static readonly string _segmentosMT = "SegmentosMT.dss";
        private static readonly string _coordMT = "CoordMT.csv";        

        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoMT;
        private StringBuilder _arqCoord;
        private readonly ModeloSDEE _SDEE;
        
        public  Param _par;

        public SegmentoMT(string alim, SqlConnectionStringBuilder connBuilder, ModeloSDEE sdee, Param par)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
            _SDEE = sdee;
        }

        //modelo
        //new line.TR1113 bus1=BMT1575B.1.2.3,bus2=BMT1568B.1.2.3,Phases=3,Linecode=CAB103_3_3,Length=0.038482,Units=km
        public bool ConsultaStoredSegmentoMT(bool _modoReconf)
        {
            _arqSegmentoMT = new StringBuilder();
            _arqCoord = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodSegmMT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km,"
                    + "CoordPAC1_x,CoordPAC1_y,CoordPAC2_x,CoordPAC2_y from dbo.StoredSegmentoMT";

                    if (_modoReconf)
                    {
                        command.CommandText  += " where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);                        
                    }
                    else
                    {
                        command.CommandText  +=  " where CodBase=@codbase and CodAlim=@CodAlim";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                        command.Parameters.AddWithValue("@CodAlim", _alim);                    
                    }
                    using ( var rs = command.ExecuteReader() )
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

                            string linha = "new line." + "TR" + rs["CodSegmMT"] //OBS1:
                                + " bus1=" + "BMT" + rs["CodPonAcopl1"] + fases //OBS1:
                                + ",bus2=" + "BMT" + rs["CodPonAcopl2"] + fases //OBS1:
                                + ",Phases=" + numFases
                                + ",Linecode=" + rs["CodCond"]
                                + ",Length=" + rs["Comp_km"]
                                + ",Units=km" + Environment.NewLine; 

                            _arqSegmentoMT.Append(linha);

                            string linhaPAC1 = "BMT" + rs["CodPonAcopl1"] + "," + rs["CoordPAC1_x"] + "," + rs["CoordPAC1_y"] + Environment.NewLine;
                            string linhaPAC2 = "BMT" + rs["CodPonAcopl2"] + "," + rs["CoordPAC2_x"] + "," + rs["CoordPAC2_y"] + Environment.NewLine;

                            _arqCoord.Append(linhaPAC1);
                            _arqCoord.Append(linhaPAC2);
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
            return _par._pathAlim + _alim + _segmentosMT;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo( _arqSegmentoMT.ToString(), GetNomeArq());
        }

        internal void GravaArqCoord()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArqCoord());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqCoord.ToString(), GetNomeArqCoord());
        }

        private string GetNomeArqCoord()
        {
            return _par._pathAlim + _alim + _coordMT;
        }

        // Une 
        internal bool UneSE(bool _modoReconfiguracao)
        {
            //condicao de saida 
            if (!_modoReconfiguracao)
                return false;

            List<string> dicCodPonAcopl = new List<string>();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodAlim,CodPonAcopl from Storedcircmt where codbase = @codbase "
                                        + "and codalim in ("+ _par._conjAlim  + ")";
                    command.Parameters.AddWithValue("@codbase", _par._codBase);

                    using ( var rs = command.ExecuteReader() )
                    {
                        // verifica ocorrencia de elemento no banco
                        if (!rs.HasRows)
                        {
                            return false;
                        }

                        // Preenche lista com os 2 cod acoplamento
                        while (rs.Read())
                        {
                            dicCodPonAcopl.Add( rs["CodPonAcopl"].ToString() );
                        }
                    }
                }

                //fecha conexao
                conn.Close();
            }

            string linha="";
                        
            // pega os pels 
            for (int i = 0; i < dicCodPonAcopl.Count;i++ )
            {
                // tratamento do ultimo elemento
                if (i == dicCodPonAcopl.Count - 1)
                {
                    linha += "new line." + "TR" + "fic" + i.ToString()
                        + " bus1=" + "BMT" + dicCodPonAcopl[i] + ".1.2.3"
                        + ",bus2=" + "BMT" + "FIC" + ".1.2.3"
                        + ",Phases=" + "3"
                        + ",Linecode=" + "CAB"
                        + ",Length=" + "0.001"
                        + ",Units=km" + Environment.NewLine;

                    _par._trEM = "TR" + "fic" + i.ToString();
                }
                else 
                {
                    linha += "new line." + "TR" + "fic" + i.ToString()
                       + " bus1=" + "BMT" + dicCodPonAcopl[i] + ".1.2.3"
                       + ",bus2=" + "BMT" + dicCodPonAcopl[i+1] + ".1.2.3"
                       + ",Phases=" + "3"
                       + ",Linecode=" + "CAB"
                       + ",Length=" + "0.001"
                       + ",Units=km" + Environment.NewLine;
                }
            }
            _arqSegmentoMT.Append(linha);

            return true;
        }
    }
}
