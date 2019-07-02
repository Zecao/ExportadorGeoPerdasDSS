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
        private static string _segmentosMT = "SegmentosMT.dss";

        private string _codBase;
        private string _pathAlim;
        private string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoMT;
        private ModeloSDEE _SDEE;

        public SegmentoMT(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase, ModeloSDEE sdee)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;
            _codBase = codBase;
            _SDEE = sdee;
        }

        //modelo
        //new line.TR1113 bus1=BMT1575B.1.2.3,bus2=BMT1568B.1.2.3,Phases=3,Linecode=CAB103_3_3,Length=0.038482,Units=km
        public bool ConsultaStoredSegmentoMT()
        {
            _arqSegmentoMT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodSegmMT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km from dbo.StoredSegmentoMT where CodBase=@codbase and CodAlim=@CodAlim";
                    command.Parameters.AddWithValue("@codbase", _codBase);
                    command.Parameters.AddWithValue("@CodAlim", _alim); 

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

                            string linha = "new line." + rs["CodSegmMT"]
                                + " bus1=" + rs["CodPonAcopl1"] + fases //OBS1:
                                + ",bus2=" + rs["CodPonAcopl2"] + fases //OBS1:
                                + ",Phases=" + numFases
                                + ",Linecode=" + rs["CodCond"]
                                + ",Length=" + rs["Comp_km"]
                                + ",Units=km" + Environment.NewLine; 

                            _arqSegmentoMT.Append(linha);
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
            return _pathAlim + _alim + _segmentosMT;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo( _arqSegmentoMT.ToString(), GetNomeArq());
        }
    }
}
