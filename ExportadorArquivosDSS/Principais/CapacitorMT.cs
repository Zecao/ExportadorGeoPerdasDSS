using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class CapacitorMT
    {
        // membros privador
        private static readonly string _capacitorMT = "CapacitorMT.dss";
        private StringBuilder _arqCapacitor;

        private readonly string _codBase;
        private readonly string _pathAlim;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;

        public CapacitorMT(string pathAlim, string alim, SqlConnectionStringBuilder connBuilder, string codBase)
        {
            _pathAlim = pathAlim;
            _alim = alim;
            _connBuilder = connBuilder;
            _codBase = codBase;
        }

        // Modelo
        // new capacitor.CAP74563,Phases=3,bus1=BMT156066088.1.2.3.0,conn=wye,Kvar=300,Kv=13.8
        public bool ConsultaBanco()
        {
            _arqCapacitor = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // TODO ADICIONAR TnsLnh1_kV	
                    command.CommandText = "select CodCapMT,CodPonAcopl,CodFas,PotNom_kVA " +
                        "from dbo.CemigCapacitorMT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string fasesDSS = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());

                            string linha = "new capacitor." + rs["CodCapMT"].ToString() 
                                + " bus1=" + "BMT" + rs["CodPonAcopl"].ToString() + fasesDSS + ".0" // OBS: o ".0" transforma em ligacao Y //OBS1
                                + ",Phases=" + numFases
                                + ",Conn=wye" 
                                + ",Kvar=" + rs["PotNom_kVA"].ToString()
                                + ",Kv=" + "13.8" + Environment.NewLine;

                            _arqCapacitor.Append(linha);
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
            return _pathAlim + _alim + _capacitorMT;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqCapacitor.ToString(), GetNomeArq());
        }
    }
}
