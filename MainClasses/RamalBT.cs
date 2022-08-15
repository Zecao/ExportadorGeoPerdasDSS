using System;
using System.Data.SqlClient;
using System.Text;


namespace ExportadorGeoPerdasDSS
{
    class RamalBT
    {
        // membros privados
        private static readonly string _ramais = "Ramais.dss";
        private Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private StringBuilder _arqSegmentoBT;

        public RamalBT(string alim, SqlConnectionStringBuilder connBuilder, Param par)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
        }

        //modelo
        // new line.RML10041834 bus1=BBT2049281.1.0,bus2=R10041834.1.0,Phases=1,Linecode=CABBTG01_4_1,Length=0.017,Units=km
        //CodBase	CodRmlBT	CodAlim	CodTrafo	CodPonAcopl1	CodPonAcopl2	CodFas	CodCond	Comp_km	Descr	CodSubAtrib	CodAlimAtrib	CodTrafoAtrib	Ordm	De	Para
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqSegmentoBT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    // se modo reconfiguracao 
                    if (_modoReconf)
                    {
                        command.CommandText = "select CodRmlBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km,Descr from dbo.StoredRamalBT "
                            +"where CodBase=@codbase and CodAlim in (" + _par._conjAlim + ")";
                        command.Parameters.AddWithValue("@codbase", _par._codBase);
                    }
                    else
                    {
                        command.CommandText = "select CodRmlBT,CodPonAcopl1,CodPonAcopl2,CodFas,CodCond,Comp_km,Descr from dbo.StoredRamalBT "
                            +"where CodBase=@codbase and CodAlim=@CodAlim";
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
                            // comprimento ramal interno
                            string compRmlInt = rs["Descr"].ToString();

                            string compRml = rs["Comp_km"].ToString();
                            
                            /* // TODO: aumento de 3metros do ramal
                            double compRml_D = Double.Parse(compRml) + 0.003;

                            compRml = compRml_D.ToString();*/
                            
                            // atribui compRml interno ao 
                            if (!compRmlInt.Equals(""))
                            {
                                compRml = compRmlInt;
                            }

                            string fases = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string linha = "new line.RML" + rs["CodRmlBT"].ToString()
                                + " bus1=" + "BBT" + rs["CodPonAcopl1"] + fases //OBS1
                                + ",bus2=" + "RML" +rs["CodPonAcopl2"] + fases //OBS1
                                + ",Phases=" + numFases
                                + ",Linecode=" + rs["CodCond"].ToString()
                                + ",Length=" + compRml
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
            return _par._pathAlim + _alim + _ramais;
        }

        internal void GravaEmArquivo()
        {
            ArqManip.SafeDelete(GetNomeArq());

            ArqManip.GravaEmArquivo(_arqSegmentoBT.ToString(), GetNomeArq());
        }
    }
}
