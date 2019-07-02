using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class SegmentoBT
    {
        // membros privados
        private static readonly string _segmentosBT = "SegmentosBT.dss";

        private readonly string _codBase;
        private readonly string _pathAlim;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoBT;

        public SegmentoBT(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;
            _codBase = codBase;
        }

        //modelo
        // new line.TBT10260958 bus1=BBT2593020.1.2.3.0,bus2=BBT2593027.1.2.3.0,Phases=3,Linecode=BT107_BT107_7_3_1,Length=0.038957,Units=km
        // CodBase	CodSegmBT	CodAlim	CodTrafo	CodPonAcopl1	CodPonAcopl2	CodFas	CodCond	Comp_km	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	Ordm	De	Para
        public bool ConsultaBanco()
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodSegmBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km from dbo.StoredSegmentoBT where CodBase=@codbase and CodAlim=@CodAlim";
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

                            string linha = "new line." + rs["CodSegmBT"].ToString()
                                + " bus1=" + "BBT" + rs["CodPonAcopl1"] + fases  //OBBS1
                                + ",bus2=" + "BBT" + rs["CodPonAcopl2"] + fases  //OBBS1
                                + ",Phases=" + numFases 
                                + ",Linecode=" + rs["CodCond"].ToString() 
                                + ",Length=" + rs["Comp_km"].ToString()
                                + ",Units=km" + Environment.NewLine;

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
            return _pathAlim + _alim + _segmentosBT;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
