using ConsoleApplication2.Principais;
using ExecutorOpenDSS;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportadorArqDSS
{
    class ExportadorArqDSS
    {
        // membros publicos -> parametros configuraveis pelo usuario

        // arquivo txt com lista de alimentadores
        private static string arqLstAlimentadores = "lstAlimentadores.m"; 

        // arquivo do Excel com somatorio em PU das curvas de carga
        public static readonly string _arqConsumoMensalPU = "somaCurvasCargaPU.xlsx";

        // prefixo arquivo .txt de feriados
        public static string _feriado = "Feriado"; //arquivo de feriado

        // path do APP
        public static readonly string _path = @"F:\DropboxZecao\Dropbox2018\Dropbox\Dropbox\0ExecutorOpenDSS\FeederExample\0_alimTese\t1\";

        // sub diretorio recursos permanentes
        public static string _permRes = "0PermRes\\";

        // servidor SGBD
        public static string _dataSource = @"NOVA\SQLEXPRESS"; //@"sa-corp-sql01\p"; 
        public static string _banco = "GEOPERDAS_2018_2";

        //Modelos SDEE 
        //Modelo GeoPerdas ANEEL
        private static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: false, modeloCarga: "ANEEL");
        //Modelo tese
        //private static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: true, utilizarCurvaDeCargaClienteMTIndividual: true, incluirCapacitoresMT: true, modeloCarga: "ANEEL");

        // mes e ano para a geracao dos arquivos de carga BT e MT
        public static int _iMes = 1;
        public static string _ano = "2018";
        public static readonly string _codBase = "4950";

        // FIM membros publicos -> parametros configuraveis usuario

        //membros privados
        private static string _alim;
        private static SqlConnectionStringBuilder _connBuilder;        
        private static string _pathAlim;

        private static bool _temSegmentoMT = false;
        private static bool _temRegulador = false;
        private static bool _temTransformador = false;
        private static bool _temChaveMT = false;
        private static bool _temCapacitor = false;
        private static bool _temRamal = false;
        private static bool _temSegmentoBT = false;
        private static bool _temCargaMT = false;
        private static bool _temCargaBT = false;
        private static bool _temGeradorMT = false;
        private static bool _temGeradorBT = false;

        // obtem Lista com numero de feriados mes X Mes
        private static List<List<int>> _numDiasFeriadoXMes;

        // utilizado por CargaMT e CargaBT
        private static Dictionary<string, double> _somaCurvaCargaDiariaPU;

        static void Main(string[] args)
        {
            // parametros banco de dados
            _connBuilder = new SqlConnectionStringBuilder();
            _connBuilder.DataSource = _dataSource;
            _connBuilder.InitialCatalog = _banco;
            _connBuilder.IntegratedSecurity = true;

            // variaveis auxiliares
            CarregaVariaveisAux();

            // lista de alimentadores
            List<string> lstAlim = AlimentadoresCemig.getTodos( GetNomeArqLstAlimentadores() );

            // para cada alimentador da lista
            foreach (string alim in lstAlim)
            {
                CriaArquivosDSS(alim);
            }

            Console.Write("Fim!");
            Console.ReadKey();
        }

        private static string GetNomeArqLstAlimentadores()
        {
            return _path + _permRes + arqLstAlimentadores;
        }

        // cria arquivos DSS
        private static void CriaArquivosDSS(string alim)
        {
            _alim = alim;

            // prenche
            _pathAlim = _path + alim + "\\";

            //Cria o diretório do alimentador, caso não exista
            if (!System.IO.Directory.Exists(_pathAlim))
            {
                System.IO.Directory.CreateDirectory(_pathAlim);
            }

            // Segmento MT
            CriaSegmentoMTDSS();

            // se alimentador nao tem segmento MT aborta
            if (!_temSegmentoMT)
            {
                Console.Write("Alimentador sem segmento MT. Abortando!");
                Console.ReadKey();
                return;
            }

            // Regulador MT
            CriaReguladorMTDSS();

            // Chave MT
            CriaChaveMT();

            // Transformador MT
            CriaTransformadorMTMTMTBTDSS();

            // Capacitor
            if ( _SDEE._incluirCapacitoresMT )
            {
                CriaCapacitorMTDSS();
            }

            // Segmento BT
            CriaSegmentoBTDSS();

            // Ramais 
            CriaRamaisDSS();

            // Carga MT
            CriaCargaMTDSS();

            // Carga BT
            CriaCargaBTDSS();

            // arquivo cabecalho
            CriaCabecalhoDSS();
        }

        private static void CarregaVariaveisAux()
        {
            // obtem Lista com numero de feriados mes X Mes
            _numDiasFeriadoXMes = AuxFunc.Feriados(_ano, GetNomeArqFeriado());         

            // preenche Dic de soma Carga Mensal - Utilizado por CargaMT e CargaBT
            _somaCurvaCargaDiariaPU = LeXLSX.XLSX2Dictionary(GetNomeArqConsumoMensalPU()); 
        }
                
        private static string GetNomeArqConsumoMensalPU()
        {
            return _path + _permRes +_arqConsumoMensalPU;
        }

        // nome arquivo feriado
        private static string GetNomeArqFeriado()
        {
            return _path + _permRes + _feriado + _ano + ".txt";
        }

        private static void CriaChaveMT()
        {
            ChaveMT oChaveMT = new ChaveMT(_pathAlim, _alim, _connBuilder, _codBase);

            // realiza consulta 
            _temChaveMT = oChaveMT.ConsultaBanco();

            // _temChaveMT
            if (_temChaveMT)
            {
                oChaveMT.GravaEmArquivo();
            }
        }

        private static void CriaReguladorMTDSS()
        {
            Regulador oRT = new Regulador(_pathAlim, _alim, _connBuilder, _codBase);

            // realiza consulta StoredReguladorMT 
            _temRegulador = oRT.ConsultaStoredReguladorMT();

            // _temRegulador
            if (_temRegulador)
            {
                oRT.GravaEmArquivo();
            }
        }

        // cria arquivo dss de segmentos de MT
        private static void CriaSegmentoMTDSS()
        {
            SegmentoMT oSegMT = new SegmentoMT(_pathAlim, _alim, _connBuilder, _codBase, _SDEE );

            // realiza consulta StoredSegmentoMT 
            _temSegmentoMT = oSegMT.ConsultaStoredSegmentoMT();

            // _temSegmentoMT
            if (_temSegmentoMT)
            {
                oSegMT.GravaEmArquivo();
            }
        }

        private static void CriaTransformadorMTMTMTBTDSS()
        {
            Trafo oTrafo = new Trafo(_pathAlim, _alim, _connBuilder, _codBase);

            // realiza consulta StoredReguladorMT 
            _temTransformador = oTrafo.ConsultaBanco();

            if (_temTransformador)
            {
                oTrafo.GravaEmArquivo();
            }
        }

        private static void CriaSegmentoBTDSS()
        {
            SegmentoBT oSegBT = new SegmentoBT(_pathAlim, _alim, _connBuilder, _codBase);

            // realiza consulta 
            _temSegmentoBT = oSegBT.ConsultaBanco();

            if (_temSegmentoBT)
            {
                oSegBT.GravaEmArquivo();
            }
        }

        private static void CriaRamaisDSS()
        {
            RamalBT oRamal = new RamalBT(_pathAlim, _alim, _connBuilder, _codBase);

            // realiza consulta 
            _temRamal = oRamal.ConsultaBanco();

            if (_temRamal)
            {
                oRamal.GravaEmArquivo();
            }
        } 

        private static void CriaCargaMTDSS()
        {
            CargaMT oCargaMT = new CargaMT(_pathAlim, _alim, _connBuilder, _codBase, _iMes, _ano, _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU, _SDEE);

            // realiza consulta 
            _temCargaMT = oCargaMT.ConsultaBanco();

            if (_temCargaMT)
            {
                oCargaMT.GravaEmArquivo();
            }
        }

        private static void CriaCargaBTDSS()
        {
            CargaBT oCargaBT = new CargaBT(_pathAlim, _alim, _connBuilder, _codBase, _iMes, _ano, _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU, _SDEE);

            // realiza consulta 
            _temCargaBT = oCargaBT.ConsultaBanco();

            if (_temCargaBT)
            {
                oCargaBT.GravaEmArquivo();
            }
        }

        // TODO falta implementar
        private static void CriaCapacitorMTDSS()
        {
            CapacitorMT oCap = new CapacitorMT(_pathAlim, _alim, _connBuilder, _codBase);

            // realiza consulta 
            _temCargaBT = oCap.ConsultaBanco();

            if (_temCargaBT)
            {
                oCap.GravaEmArquivo();
            }
        }

        // modelo
        // TenNom_kV TenOpe_pu   CodPonAcopl
        /*
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
        public static List<string> ConsultaStoredCircMT()
        {
            string linha="";

            string strCab = "";

            string strEM = "";
            
            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                conn.Open();
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "select TenNom_kV,TenOpe_pu,CodPonAcopl from dbo.storedcircmt where CodBase=@codbase and CodAlim=@CodAlim";
                    command.Parameters.AddWithValue("@codbase", _codBase);
                    command.Parameters.AddWithValue("@CodAlim", _alim );

                    using (var rs = command.ExecuteReader())
                    {
                        rs.Read();

                        // verifica se NAO tem linhas
                        if (!rs.HasRows)
                        {
                            Console.Write("Erro! Alimentador não localizado na StoredCircMT");
                            return null;
                            }
                        else
                        {                            
                            linha = CriaStrCabecalho(rs);

                            strEM = CriaStrEnergyMeter(rs);

                            //adiciona 
                            linha += strEM;

                            strCab = CriaStrCabecalhoCustomizacao(rs);
                        }         
                   }
                }
            }
            // concatena resultados
            return new List<string> { linha, strCab, strEM };
        }

        // Cria Str Energymeter e comandos adicionais
        private static string CriaStrEnergyMeter(SqlDataReader rs)
        {
            string linha = "";

            // energymeter
            if (_temSegmentoMT)
            {
                linha = Environment.NewLine + "new energymeter.carga element=line." + GetTrechoEnergyMeter(rs["CodPonAcopl"].ToString())
                    + ",terminal=1" + Environment.NewLine;
            }
            // voltage bases
            linha += Environment.NewLine + "Set voltagebases=[" + rs["TenNom_kV"].ToString() + " 0.24 0.22]" + Environment.NewLine;

            // CalcVoltageBases
            linha += Environment.NewLine + "CalcVoltageBases" + Environment.NewLine;

            // 
            linha += Environment.NewLine + "! Solve mode=daily,hour=0,number=24,stepsize=1h";

            return linha;
        }

        private static string CriaStrCabecalhoCustomizacao(SqlDataReader rs)
        {
            string linha = "! manter este comentario para compatibilidade do C#" + Environment.NewLine;

            // modelo
            // new circuit.alimAFNU16 bus1=BMT155673172,basekv=13.8,pu=
            linha += "new circuit.alim" + _alim
            + " bus1=" + rs["CodPonAcopl"].ToString() + ".1.2.3"
            + ",basekv=" + rs["TenNom_kV"].ToString()
            + ",pu=" + rs["TenOpe_pu"].ToString();

            return linha;
        }

        private static string CriaStrCabecalho(SqlDataReader rs)
        {
            string linha = "";

            // cabeca alim
            string cabAlim = rs["CodPonAcopl"].ToString();

            // limpa
            linha = "clear" + Environment.NewLine;

            // new circuit
            linha += "new circuit.alim" + _alim
                + " bus1=" + cabAlim + ".1.2.3"
                + ",basekv=" + rs["TenNom_kV"].ToString()
                + ",pu=" + rs["TenOpe_pu"].ToString() + Environment.NewLine + Environment.NewLine;

            // curva de carga 
            //TODO NO momento o programa nao gera arquivo de curvas de carga
            linha += "Redirect ..\\0PermRes\\NovasCurvasTxt\\CurvasDeCargaDU.dss" + Environment.NewLine;

            // arquivo condutores 
            //TODO No momento o programa na gera o arquivo de condutores.
            linha += "Redirect ..\\0PermRes\\Condutores.dss" + Environment.NewLine;

            if (_temSegmentoMT)
            {
                linha += "Redirect " + _alim + "SegmentosMT.dss" + Environment.NewLine;
            }
            if (_temChaveMT)
            {
                linha += "Redirect " + _alim + "ChavesMT.dss" + Environment.NewLine;
            }
            if (_temRegulador)
            {
                linha += "Redirect " + _alim + "Reguladores.dss" + Environment.NewLine;
            }
            if (_temTransformador)
            {
                linha += "Redirect " + _alim + "Transformadores.dss" + Environment.NewLine;
            }
            if (_temSegmentoBT)
            {
                linha += "Redirect " + _alim + "SegmentosBT.dss" + Environment.NewLine;
            }
            if (_temRamal)
            {
                linha += "Redirect " + _alim + "Ramais.dss" + Environment.NewLine;
            }

            // TODO gerar carga para todos os anos
            if (_temCargaMT)
            {
                linha += "Redirect " + _alim + "CargaMT_" + AuxFunc.intMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }
            // TODO gerar carga para todos os anos
            if (_temCargaBT)
            {
                linha += "Redirect " + _alim + "CargaBT_" + AuxFunc.intMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }
            // OBS: comentado
            if (_temCapacitor)
            {
                linha += "! Redirect " + _alim + "CapacitorMT" + AuxFunc.intMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }
            // TODO implementar
            if (_temGeradorMT)
            {
                linha += "Redirect " + _alim + "GeradorMT" + AuxFunc.intMes2strMes(_iMes) + ".dss" + Environment.NewLine;
            }
            // TODO implementar
            if (_temGeradorBT)
            {
                linha += "Redirect " + _alim + "GeradorBT" + AuxFunc.intMes2strMes(_iMes) + ".dss" + Environment.NewLine;
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
                    command.Parameters.AddWithValue("@codbase", _codBase);
                    command.Parameters.AddWithValue("@CodAlim", _alim); 
                    command.Parameters.AddWithValue("@CodPonAcopl1", pel);

                    using (var rs = command.ExecuteReader())
                    {
                        rs.Read();

                        trecho = rs["CodSegmMT"].ToString();
                    }
                }
            }
            return trecho;
        }

        private static void CriaCabecalhoDSS()
        {
            List<string> infoCabecalho = ConsultaStoredCircMT();

            if (infoCabecalho != null)
            { 
                // grava arquivo para ser utilizado pela OpenDSS 
                ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArqCabecalho());

                ExecutorOpenDSS.ArqManip.GravaEmArquivo(infoCabecalho[0], GetNomeArqCabecalho());

                // arquivo para ser utilizado pela customizacao COM do OpenDSS
                ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArqCabecalhoCOM());

                ExecutorOpenDSS.ArqManip.GravaEmArquivo(infoCabecalho[1], GetNomeArqCabecalhoCOM());

                // arquivo para ser utilizado pela customizacao COM do OpenDSS
                ExecutorOpenDSS.ArqManip.SafeDelete(GetNomeArquivoB());

                ExecutorOpenDSS.ArqManip.GravaEmArquivo(infoCabecalho[2], GetNomeArquivoB());
            }
        }

        private static string GetNomeArquivoB()
        {
            return _pathAlim + _alim + "AnualB.dss";
        }

        private static string GetNomeArqCabecalhoCOM()
        {
            return _pathAlim + _alim + ".dss";
        }

        private static string GetNomeArqCabecalho()
        {
            return _pathAlim + _alim + "AnualA.dss";
        }
    }
}

