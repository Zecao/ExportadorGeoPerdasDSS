using System;
using System.Data.SqlClient;
using System.Text;

namespace ExportadorGeoPerdasDSS
{
    class GeradorMT
    {
        // membros privador
        private static readonly string _geradorMT = "GeradorMT_";
        private static int _iMes;

        private StringBuilder _arqGeradorMT;
        private readonly Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;

        public GeradorMT(string alim, SqlConnectionStringBuilder connBuilder, Param par, int iMes)
        {
            _iMes = iMes;
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
        }


        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqGeradorMT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // 
                    command.CommandText = "select CodGeraMT,CodAlim,CodFas,CodPonAcopl,TnsLnh_kV," +
                        "EnerMedid01_MWh,EnerMedid02_MWh,EnerMedid03_MWh,EnerMedid04_MWh,EnerMedid05_MWh,EnerMedid06_MWh," +
                        "EnerMedid07_MWh,EnerMedid08_MWh,EnerMedid09_MWh,EnerMedid10_MWh,EnerMedid11_MWh,EnerMedid12_MWh,Descr ";
                        
                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText += "from dbo.StoredGeradorMT where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);

                    }
                    else
                    {
                        command.CommandText += "from dbo.StoredGeradorMT where CodBase=@codbase and CodAlim=@CodAlim";
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
                            string CodGeraMT = rs["CodGeraMT"].ToString();

                            // curva PU
                            string linha = "new loadshape.c" + CodGeraMT + " npts=24,interval=1.0,mult=" + rs["Descr"].ToString() + Environment.NewLine;

                            // Obtem a geracao de acordo com o mes
                            string geracaoMes = AuxFunc.GetConsumoMesCorrente(rs, _iMes);

                            // usina
                            linha += "new generator." + CodGeraMT
                            + " bus1=" + "BMT" + rs["CodPonAcopl"] + ".1.2.3"
                            + ",Phases=3"
                            + ",kv=" + rs["TnsLnh_kV"].ToString()
                            + ",kW=" + geracaoMes
                            + ",pf=0.92"
                            + ",model=1" 
                            + ",daily=c" + CodGeraMT
                            + ",status=Variable" + Environment.NewLine;

                            _arqGeradorMT.Append(linha);
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
            string strMes = AuxFunc.IntMes2strMes(_iMes);

            return _par._pathAlim + _alim + _geradorMT + strMes + ".dss";
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ArqManip.GravaEmArquivo(_arqGeradorMT.ToString(), GetNomeArq());
        }
    }
}
