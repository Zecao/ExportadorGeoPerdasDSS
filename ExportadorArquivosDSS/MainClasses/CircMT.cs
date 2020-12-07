using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ConsoleApplication2.Principais
{
    class CircMT
    {
        private static Param _par;
        private static string _alim;
        private static SqlConnectionStringBuilder _connBuilder;
        private static bool _modoReconf;
        private static bool _criaArqCoordenadas;
        private static StrBoolElementosSDE _structElem;
        private static int _iMes;
        private static string _coordMT;
        private static bool _alimAtipico;

        public CircMT(string alim, SqlConnectionStringBuilder connBuilder, Param par, bool modoReconfiguracao, 
            bool criaArqCoordenadas, StrBoolElementosSDE structElem, int iMes, string coordMT, bool alimAtipico)
        {            
            _par = par;
            _alim = alim;
            _connBuilder = connBuilder;
            _modoReconf = modoReconfiguracao;
            _criaArqCoordenadas = criaArqCoordenadas;
            _structElem = structElem;
            _iMes = iMes;
            _coordMT = coordMT;
            _alimAtipico = alimAtipico;
        }

        private static string GetNomeArqCoord()
        {
            return _alim + _coordMT;
        }

        /* // estrutura tipica do arquivo masterDSS criado
        clear

        new circuit.alimAFNU16 bus1=BMT155673172,basekv=13.8,pu=1.036

        ! Arquivo de curvas de carga normalizadas
        Redirect "..\PermRes\NovasCurvasTxt\CurvasDeCargaDU.dss"

        ! Arquivo de condutores
        Redirect "..\PermRes\Condutores.dss"

        Redirect AFNU16Transformadores.dss
        Redirect AFNU16Ramais.dss
        Redirect AFNU16ChavesMT.dss
        Redirect AFNU16SegmentosBT.dss
        Redirect AFNU16SegmentosMT.dss
        Redirect AFNU16Carga_BT.dss
        Redirect AFNU16Carga_MT.dss

        new energymeter.carga element=line.TR12301107,terminal=1

        ! Seta as tensoes de base do sistema
        Set voltagebases="13.8 0.24 0.22"

        ! Calcula as tensoes das linhas em pu
        CalcVoltageBases

        !Set mode=daily  hour=0 number=24 stepsize=1h

        Solve mode=daily  hour=0 number=24 stepsize=1h
        */
        public List<string> ConsultaStoredCircMT()
        {
            string stringMasterDSS_A = "";

            string strCab = "";

            string stringMasterDSSP_B = "";

            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                conn.Open();
                using (SqlCommand command = conn.CreateCommand())
                {
                    // consulta
                    command.CommandText = "select TenNom_kV,TenOpe_pu,CodPonAcopl from dbo.storedcircmt ";             

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
                        rs.Read();

                        // verifica se NAO tem linhas
                        if (!rs.HasRows)
                        {
                            Console.Write(_alim + ": não localizado na StoredCircMT");
                            return null;
                        }
                        else
                        {
                            // cria String MasterDSS Parte A
                            stringMasterDSS_A = CriaStringMasterDSS_ParteA(rs);

                            // cria String MasterDSS Parte B
                            stringMasterDSSP_B = CriaStringMasterDSS_ParteB(rs);

                            // OBS: As partes A e B são adicionadas e posteriormente gravadas e um único arquivo 
                            // com a seguinte nomenclaura XXXAnualA.dss, onde XXX eh o nome do alimentador
                            // OBS: A parte B eh gravada com o nome XXXAnualB eh utilizada pela customização "ExecutorOpenDSS"
                            stringMasterDSS_A += stringMasterDSSP_B;

                            // cria string arquivo MasterDSS que sera utilizada pela customizacao "ExecutorOpenDSS"
                            // com a seguinte nomenclaura XXX.dss, onde XXX eh o nome do alimentador
                            strCab = CriaStrCabecalhoCustomizacao(rs);
                        }
                    }
                }
            }
            // concatena resultados
            return new List<string> { stringMasterDSS_A, strCab, stringMasterDSSP_B };
        }

        // Cria Str Energymeter e comandos adicionais
        private static string CriaStringMasterDSS_ParteB(SqlDataReader rs)
        {
            string linha = "";

            if (_modoReconf)
            {
                // OBS: aloca medidor no terminal 2
                linha = Environment.NewLine + "new energymeter.carga element=line." + _par._trEM
                    + ",terminal=2" + Environment.NewLine; //OBS1
            }
            else
            {
                linha = Environment.NewLine + "new energymeter.carga element=line." + "TR" + GetTrechoEnergyMeter(rs["CodPonAcopl"].ToString())
                    + ",terminal=1" + Environment.NewLine; //OBS1
            }

            // voltage bases
            linha += Environment.NewLine + "Set voltagebases=[" + rs["TenNom_kV"].ToString() + " 0.24 0.22]" + Environment.NewLine;

            // CalcVoltageBases
            linha += "CalcVoltageBases" + Environment.NewLine;

            // 
            linha += Environment.NewLine + "! Solve mode=daily,hour=0,number=24,stepsize=1h";

            if (_criaArqCoordenadas)
            {
                // Linhas de coordenadas
                linha += Environment.NewLine + Environment.NewLine + "! BusCoords " + GetNomeArqCoord();
            }
            return linha;
        }

        private static string CriaStrCabecalhoCustomizacao(SqlDataReader rs)
        {
            string linha = "! manter este comentario para compatibilidade do C#" + Environment.NewLine;

            // add linha circuito
            linha += CriaStrCircuit(rs);

            return linha;
        }

        private static string CriaStrCircuit(SqlDataReader rs)
        {
            if (_modoReconf)
            {
                return "new circuit.alim" + _alim
                + " bus1=" + "BMT" + "FIC" + ".1.2.3"
                + ",basekv=" + rs["TenNom_kV"].ToString()
                + ",pu=" + rs["TenOpe_pu"].ToString()
                + Environment.NewLine + Environment.NewLine;
            }
            else
            {
                // modelo
                // new circuit.alimAFNU16 bus1=BMT155673172,basekv=13.8,pu=
                return "new circuit.alim" + _alim
                + " bus1=" + "BMT" + rs["CodPonAcopl"].ToString() + ".1.2.3" //OBS1
                + ",basekv=" + rs["TenNom_kV"].ToString()
                + ",pu=" + rs["TenOpe_pu"].ToString()
                + Environment.NewLine + Environment.NewLine; 
            }
        }

        // funcao que cria a string do arquivo masterDSS
        private static string CriaStringMasterDSS_ParteA(SqlDataReader rs)
        {
            string linha = "";

            // cabeca alim
            string cabAlim = rs["CodPonAcopl"].ToString();

            // limpa
            linha = "clear" + Environment.NewLine;

            // new circuit
            linha += CriaStrCircuit(rs);

            // curva de carga 
            //TODO NO momento o programa nao gera arquivo de curvas de carga
            linha += "Redirect ..\\0PermRes\\NovasCurvasTxt\\CurvasDeCargaDU.dss" + Environment.NewLine;

            // arquivo condutores 
            //TODO No momento o programa na gera o arquivo de condutores.
            linha += "Redirect ..\\0PermRes\\Condutores.dss" + Environment.NewLine;

            if (_structElem._temSegmentoMT)
            {
                linha += "Redirect " + _alim + "SegmentosMT.dss" + Environment.NewLine;
            }
            if (_structElem._temChaveMT)
            {
                linha += "Redirect " + _alim + "ChavesMT.dss" + Environment.NewLine;
            }
            if (_structElem._temRegulador)
            {
                linha += "Redirect " + _alim + "Reguladores.dss" + Environment.NewLine;
            }
            if (_structElem._temTransformador)
            {
                linha += "Redirect " + _alim + "Transformadores.dss" + Environment.NewLine;
            }
            if (_structElem._temSegmentoBT)
            {
                linha += "Redirect " + _alim + "SegmentosBT.dss" + Environment.NewLine;
            }
            if (_structElem._temRamal)
            {
                linha += "Redirect " + _alim + "Ramais.dss" + Environment.NewLine;
            }

            if (_structElem._temCargaMT)
            {
                linha += "Redirect " + _alim + "CargaMT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }

            if (_structElem._temCargaBT)
            {
                linha += "Redirect " + _alim + "CargaBT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }

            if (_structElem._temCapacitorMT)
            {
                linha += "Redirect " + _alim + "CapacitorMT.dss" + Environment.NewLine;
            }

            if ((_alimAtipico) & (_structElem._temGeradorMT))
            {
                linha += "Redirect " + _alim + "GeradorMT_" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }

            // TODO implementar
            if (_structElem._temGeradorBT)
            {
                linha += "Redirect " + _alim + "GeradorBT" + AuxFunc.IntMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }
            return linha;
        }

        // obtem o 1 trecho do alimentador
        private static string GetTrechoEnergyMeter(string pel)
        {
            string trecho = "";
            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                conn.Open();
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select CodSegmMT from dbo.StoredSegmentoMT where CodBase=@codbase and CodAlim=@CodAlim and CodPonAcopl1=@CodPonAcopl1";
                    command.Parameters.AddWithValue("@codbase", _par._codBase);
                    command.Parameters.AddWithValue("@CodAlim", _alim);
                    command.Parameters.AddWithValue("@CodPonAcopl1", pel);

                    using (var rs = command.ExecuteReader())
                    {
                        //
                        if (!rs.HasRows)
                        {
                            Console.Write("Nao encontrado trecho do Energymeter");
                            return trecho;
                        }

                        rs.Read();

                        trecho = rs["CodSegmMT"].ToString();
                    }
                }
            }
            return trecho;
        }

        internal void GravaEmArquivo()
        {
            throw new NotImplementedException();
        }
    }
}
