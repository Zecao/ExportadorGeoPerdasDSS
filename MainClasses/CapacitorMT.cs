using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class CapacitorMT
    {
        private static readonly string _capacitorMT = "CapacitorMT.dss";
        private StringBuilder _arqCapacitor;
        private readonly Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;

        public CapacitorMT(string alim, SqlConnectionStringBuilder connBuilder, Param par)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
        }

        // Modelo
        // new capacitor.CAP74563,Phases=3,bus1=BMT156066088.1.2.3.0,conn=wye,Kvar=300,Kv=13.8
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqCapacitor = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText = "select CodCapMT,CodPonAcopl,CodFas,PotNom_KVAr,kvnom " +
                            "from " + _par._schema + "CemigCapacitorMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = "select CodCapMT,CodPonAcopl,CodFas,PotNom_KVAr,kvnom " +
                            "from " + _par._schema + "CemigCapacitorMT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string fasesDSS = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string kvbase = rs["kvnom"].ToString();

                            string linha="";

                            // se banco trifasico
                            if (numFases.Equals("3"))
                            {
                                // calcula potencia por fase
                                string potFase = AuxFunc.GetPotPorFase(rs["PotNom_KVAr"].ToString());

                                string[] fases={"1","2","3"};

                                foreach(string fase in fases)
                                {
                                    linha += "new capacitor." + "CAP" + rs["CodCapMT"].ToString() + "-" + fase
                                       + " bus1=" + "BMT" + rs["CodPonAcopl"].ToString()
                                       + "." + fase + ".0" // OBS: o ".0" transforma em ligacao Y //OBS1
                                       + ",Phases=1"
                                       + ",Conn=LN"
                                       + ",Kvar=" + potFase
                                       + ",Kv=" + kvbase + Environment.NewLine;
                                }
                            }
                            // capacitor monofasico
                            else 
                            {
                                linha = "new capacitor." + "CAP" + rs["CodCapMT"].ToString()
                                   + " bus1=" + "BMT" + rs["CodPonAcopl"].ToString()
                                   + fasesDSS + ".0" // OBS: o ".0" transforma em ligacao Y //OBS1
                                   + ",Phases=1" 
                                   + ",Conn=LN"
                                   + ",Kvar=" + rs["PotNom_KVAr"].ToString()
                                   + ",Kv=" + kvbase + Environment.NewLine;
                            
                            }
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
            return _par._pathAlim + _alim + _capacitorMT;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqCapacitor.ToString(), GetNomeArq());
        }
    }
}
