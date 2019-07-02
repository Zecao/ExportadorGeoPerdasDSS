using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class RamalBT
    {
        // membros privados
        private static string _ramais = "Ramais.dss";

        private string _codBase;
        private string _pathAlim;
        private string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoBT;

        public RamalBT(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;
            _codBase = codBase;
        }

        //modelo
        // new line.RML10041834 bus1=BBT2049281.1.0,bus2=R10041834.1.0,Phases=1,Linecode=CABBTG01_4_1,Length=0.017,Units=km
        //CodBase	CodRmlBT	CodAlim	CodTrafo	CodPonAcopl1	CodPonAcopl2	CodFas	CodCond	Comp_km	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	Ordm	De	Para
        public bool ConsultaBanco()
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodRmlBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km from dbo.StoredRamalBT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string linha = "new line.RML" + rs["CodRmlBT"].ToString()
                                + " bus1=" + "BBT" + rs["CodPonAcopl1"] + fases //OBS1
                                + ",bus2=" + "RML" +rs["CodPonAcopl2"] + fases //OBS1
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
            return _pathAlim + _alim + _ramais;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
