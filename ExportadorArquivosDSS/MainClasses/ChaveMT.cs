using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2.Principais
{
    class ChaveMT
    {
        // membros privados
        private static readonly string _chave = "ChavesMT.dss"; 
        private StringBuilder _arqChaveMT;
        private Param _par;
        private readonly string _alim;
        private static SqlConnectionStringBuilder _connBuilder;

        public ChaveMT(string alim, SqlConnectionStringBuilder connBuilder, Param par)
        {
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
        }

        // new line.CTR100934 bus1=BMT98417853.2,bus2=BMT98417888.2,Phases=1,LineCode=tieSwitch1,Length=0.001,Units=km,switch=T
        // open line.CTR522670 term= 1
        // CodBase	CodChv	CodAlim	CodPonAcopl1	CodPonAcopl2	CodFas	EstChv	Descr	CodSubAtrib	CodAlimAtrib	Ordm	De	Para
        public bool ConsultaBanco(bool _modoReconf)
        {
            _arqChaveMT = new StringBuilder();

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodChvMT,CodPonAcopl1,CodPonAcopl2,CodFas,EstChv,Descr "
                            + "from dbo.StoredChaveMT ";

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
                            string fasesDSS = AuxFunc.GetFasesDSS(rs["CodFas"].ToString());
                            string numFases = AuxFunc.GetNumFases(rs["CodFas"].ToString());
                            string lineCode = "tieSwitch";
                            string codChave = rs["CodChvMT"].ToString();

                            //se chave mono
                            if (numFases.Equals("1"))
                            {
                                lineCode = "tieSwitch1";
                            }                           
                            
                            string linha = "new line.CTR" + codChave //OBS1 adicionei prefixo CTR
                                + " bus1=" + "BMT" + rs["CodPonAcopl1"].ToString() + fasesDSS  //OBS1
                                + ",bus2=" + "BMT" + rs["CodPonAcopl2"].ToString() + fasesDSS  //OBS1
                                + ",Phases=" + numFases
                                + ",LineCode=" + lineCode
                                + ",Length=0.001,Units=km,switch=T"
                                + " !numEQ " + rs["Descr"].ToString() + Environment.NewLine;

                            // se chave fechada 
                            if (rs["EstChv"].ToString().Equals("1"))
                            {
                                linha += "open line.CTR" + codChave + " term=1" + Environment.NewLine; //OBS1 adicionei prefixo CTR
                            }

                            // if recloser 
                            if (codChave.Contains("R")) //"line.CTRR17947
                            {
                                linha += "new Recloser." + codChave + " monitoredobj='line.CTR" + codChave + "' monitoredterm=1,numfast=1,"
                                    + "phasedelayed='VERY_INV',grounddelayed='VERY_INV',phasetrip=280,TDPhDelayed=1,GroundTrip=45,TDGrDelayed=14,"
                                    + "shots=1,recloseintervals=(10, 20, 20)" + Environment.NewLine; 
                            }

                            _arqChaveMT.Append(linha);
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
            return _par._pathAlim + _alim + _chave;
        }

        internal void GravaEmArquivo()
        {
            ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArq());

            // grava em arquivo
            ExecutorOpenDSS.ArqManip.GravaEmArquivo(_arqChaveMT.ToString(), GetNomeArq());
        }
    }
}
