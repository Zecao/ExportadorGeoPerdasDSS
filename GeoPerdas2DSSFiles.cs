using AuxClasses;
using ExportadorGeoPerdasDSS;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExportadorArqDSS
{
    class GeoPerdas2DSSFiles
    {
        // codbase
        public static string _codBase = "2022114950"; //"2022114950"; //"2021124950"; //"2020014950"; // "4950"; 
        public static string _schema = "geo2022.";//"geo2023."; //"geo2022."; //""; //"geo2021." geo2022;

        // mes e ano para a geracao dos arquivos de carga BT e MT
        public static int _iMes = 12; //   PVSYstemBT 3,4,5,6 2023
        public static string _ano = "2022"; // 2021 //2022

        public static bool _criaTodosOsMeses = false;  // flag p/ criar todos os meses de carga MT BT e geradores
        public static bool _criaArqCoordenadas = true; // flag p/ criar arq coordenadas

        /* 
        Parameters order _modelPVSystems, _invControlModeMV, _varFollowInvMV, _PVPowerFactorMV

        _modelPVSystems     -> False = exports PVSystems as generator model=1
        _geraInvControl     _> true to generate invControl
        _invControlModeMV   -> Default: "VOLTVAR" "voltwatt"
        _varFollowInvMV     -> Default: false (Have to set False to enable Inverter night mode)
        _voltVarcurve =     -> Default: "voltvar_c"
        */
        //true, "VOLTVAR", "True", "voltvar_c"
        static readonly PVSystemPar _pvMV = new PVSystemPar(true, true); // true, true, "VOLTVAR", "True", "voltvar_c"
        static readonly PVSystemPar _pvLV = new PVSystemPar(false, false);

        /* VarFollowInverter: Boolean variable which indicates that the reactive power does not respect the inverter status.
        – When set to True, PVSystem’s reactive power will cease when the inverter status is OFF,
        due to the power from PV array dropping below %cutout. The reactive power will begin
        again when the power from PV array is above %cutin;
        – When set to False, PVSystem will provide/absorb reactive power regardless of the status
        of the inverter.*/

        public static bool _criaDispProtecao = false; // flag p/ dispositivos de protecao (Recloser e Fuses) && taxas de falhas em lines
        public static bool _modelo4condutores = false; //modelo 4 condutor BT

        // cria arquivo DSS com 2 alimentadores para uso da Reconfiguracao
        public static bool _genAllSubstation = false; // Generates all substation feeders as one. Uses the first feeder for directory name.

        // arquivo txt com lista de alimentadores
        public static string _arqLstAlimentadores = "lstAlimentadores.m";

        // lista SEs (usado em reconfiguracao somente)
        public static string _arqLstSEs = "lstSEs.m";

        // arquivo do Excel com somatorio em PU das curvas de carga
        public static readonly string _arqConsumoMensalPU = "somaCurvasCargaPU.xlsx";

        // arquivo do Excel com curvas individuais dos clientes primarios
        public static readonly string _arqCurvaCargaCliMT = "curvasTipicasClientesMT_2018.xlsx";

        // prefixo arquivo .txt de feriados
        public static string _feriado = "Feriado"; //arquivo de feriado

        // path do APP - CASA
        //public static readonly string _path = @"F:\DropboxZecao\Dropbox\0doutorado\0soft\0alimCemig\";

        //CEMIG - Servidor Pertec
        //public static readonly string _path = @"\\pwnap-pertec01\OPenDSSCemig\2022_lm1\";
        public static readonly string _path = @"\\pwnap-pertec01\OPenDSSCemig\2022_CAP\";
        //public static readonly string _path = @"\\pwnap-pertec01\OPenDSSCemig\2023_2\";

        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\2022\";
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\2022_SemGDBT\";
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\2022_SemGDMTBT\";
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\2023_2\";
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\01_voltvar_PF1\";
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\01_voltvar\"; //OBS: VOLTVAR
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\01_voltvar_Noite\";

        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\OpenDSSLocal\2022_reconf\";

        // TODO FIX this hard coded 
        // sub diretorio recursos permanentes
        public static string _permRes = "0PermRes\\";

        // servidor SGBD
        public static string _banco = "GEOPERDAS_2022"; // GEOPERDAS_2022 GEOPERDAS_2021 GEOPERDAS_2020 GEOPERDAS_2019 

        // banco
        public static string _dataSource = @"PWNBS-PERTEC01\PTEC"; //@"sa-corp-sql0";

        //Modelo PADRAO (GeoPerdas ANEEL)
        //OBS: Capacitor pode ser colocado na hora da execucao
        //private static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: false, modeloCarga: "ANEEL",
        //    reatanciaTrafos: false);

        // sim3 modelo de carga PCONST
        private readonly static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: true, modeloCarga: "PCONST",
           reatanciaTrafos: true);

        // FIM membros publicos -> parametros configuraveis usuario

        //membros privados
        public static Param _par;

        private static SqlConnectionStringBuilder _connBuilder;

        private static StrBoolElementosSDE _structElem;

        // arquivo de coordenadas
        private static readonly string _coordMT = "CoordMT.csv";

        // obtem Lista com numero de feriados mes X Mes
        private static List<List<int>> _numDiasFeriadoXMes;

        // utilizado por CargaMT e CargaBT
        private static Dictionary<string, double> _somaCurvaCargaDiariaPU;

        // utilizado por CargaMT e CargaBT
        private static Dictionary<string, List<string>> _curvasTipicasClientesMT;

        static void Main() //string[] args
        {
            // parametros banco de dados
            _connBuilder = new SqlConnectionStringBuilder();
            _connBuilder.DataSource = _dataSource;
            _connBuilder.InitialCatalog = _banco;
            _connBuilder.IntegratedSecurity = false;
            _connBuilder.UserID = "U_DBPERTEC01";
            _connBuilder.Password = "";

            // variaveis auxiliares
            CarregaVariaveisAux();

            // se modo reconfiguracao
            if (_genAllSubstation)
            {
                GenSubstationDSSFiles();
            }
            else
            {
                GeneratesFeedersDSSFiles();
            }

            Console.Write("Fim!");
            Console.ReadKey();
        }

        // cria arquivos DSS
        private static void CriaArquivosDSS()
        {
            //Cria o diretório do alimentador, caso não exista
            if (!System.IO.Directory.Exists(_par._pathAlim))
            {
                System.IO.Directory.CreateDirectory(_par._pathAlim);
            }

            // Segmento MT
            CriaSegmentoMTDSS();

            // Se nao tem segmento, aborta 
            if (!_structElem._temSegmentoMT)
            {
                return;
            }

            // Regulador MT
            CriaReguladorMTDSS();

            // Chave MT
            CriaChaveMT();

            // Transformador MT
            CriaTransformadorMTMTMTBTDSS();

            // Capacitor
            if (_SDEE._incluirCapacitoresMT)
            {
                CriaCapacitorMTDSS();
            }

            CriaSegmentoBTDSS();

            // Ramais 
            CriaRamaisDSS();

            // Carga MT
            CriaCargaMTDSS();

            // Carga BT
            CriaCargaBTDSS();

            // Gerador MT
            CriaGeradorMT();

            // Gerador BT
            CriaGeradorBT();

            // arquivo cabecalho
            CriaCabecalhoDSS();
        }

        // Generates entire substation
        private static void GenSubstationDSSFiles()
        {
            // populates lstAlim with all feeders from txt file "Alimentadores.m"

            List<string> lstAlim = CemigFeeders.GetAllFeedersFromTxtFile(GetNomeArqLstAlimentadores());

            _par = new Param(_path, _permRes, _codBase, _modelo4condutores, _schema, "", _ano, _pvMV, _pvLV);

            // para cada SE da lista
            foreach (string alim in lstAlim)
            {
                // sets current feeder
                _par.SetCurrentAlim(alim);

                // gets substation
                string substation = System.Text.RegularExpressions.Regex.Replace(alim, @"[\d-]", string.Empty);

                // gets all feeders name from substation //TODO move to Param ?
                bool ret = CemigFeeders.GetAllFeedersFromSubstationString(substation, _connBuilder, _par);

                if (_par._conjAlim == null)
                {
                    Console.Write(substation + " não encontrado!\n");
                    continue;
                }

                // creates dss files
                CriaArquivosDSS();
            }
        }

        // Generates entire substation
        private static void GeneratesFeedersDSSFiles()
        {
            // lista de alimentadores
            List<string> lstAlim = CemigFeeders.GetAllFeedersFromTxtFile(GetNomeArqLstAlimentadores());

            _par = new Param(_path, _permRes, _codBase, _modelo4condutores, _schema, "", _ano, _pvMV, _pvLV);

            // para cada alimentador da lista
            foreach (string alim in lstAlim)
            {
                // sets current feeder
                _par.SetCurrentAlim(alim);

                // creates dss files
                CriaArquivosDSS();

                // TODO
                //ClassificaCurvaBT curvaBT = new ClassificaCurvaBT(_connBuilder, _par);
            }
        }

        private static string GetNomeArqLstAlimentadores()
        {
            return _path + _permRes + _arqLstAlimentadores;
        }

        private static void CriaGeradorBT()
        {
            if (_criaTodosOsMeses)
            {
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    _iMes = i;

                    //
                    CriaGeradorBTPvt();
                }
            }
            else
            {
                //
                CriaGeradorBTPvt();
            }
        }

        private static void CriaGeradorBTPvt()
        {
            GeradorBT oGerBT = new GeradorBT(_connBuilder, _par, _iMes);

            // realiza consulta  
            _structElem._temGeradorBT = oGerBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temGeradorBT)
            {
                oGerBT.GravaEmArquivo();
            }
        }

        private static void CarregaVariaveisAux()
        {
            // obtem Lista com numero de feriados mes X Mes
            _numDiasFeriadoXMes = AuxFunc.Feriados(GetNomeArqFeriado());

            // preenche Dic de soma Carga Mensal - Utilizado por CargaMT e CargaBT
            _somaCurvaCargaDiariaPU = XLSXFile.XLSX2Dictionary(GetNomeArqConsumoMensalPU());

            // preenche Dic com curvas de carga INDIVIDUAIS da CargaMT
            if (_SDEE._utilizarCurvaDeCargaClienteMTIndividual)
            {
                _curvasTipicasClientesMT = XLSXFile.XLSX2DictString(GetNomeArqCurvaCargaCliMT());
            }
        }

        private static string GetNomeArqCurvaCargaCliMT()
        {
            return _path + _permRes + _arqCurvaCargaCliMT;
        }

        private static string GetNomeArqConsumoMensalPU()
        {
            return _path + _permRes + _arqConsumoMensalPU;
        }

        // nome arquivo feriado
        private static string GetNomeArqFeriado()
        {
            return _path + _permRes + _feriado + _ano + ".txt";
        }

        private static void CriaGeradorMT()
        {
            if (_criaTodosOsMeses)
            {
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    _iMes = i;

                    //
                    CriaGeradorMTPvt();
                }
            }
            else
            {
                //
                CriaGeradorMTPvt();
            }
        }

        //
        private static void CriaGeradorMTPvt()
        {
            GeradorMT oGerMT = new GeradorMT(_connBuilder, _par, _iMes);

            // realiza consulta  
            _structElem._temGeradorMT = oGerMT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temGeradorMT)
            {
                oGerMT.GravaEmArquivo();
            }
        }

        private static void CriaChaveMT()
        {
            ChaveMT oChaveMT = new ChaveMT(_connBuilder, _par, _criaDispProtecao);

            // realiza consulta 
            _structElem._temChaveMT = oChaveMT.ConsultaBanco(_genAllSubstation);

            // _temChaveMTno 
            if (_structElem._temChaveMT)
            {
                oChaveMT.GravaEmArquivo();
            }
        }

        private static void CriaReguladorMTDSS()
        {
            Regulador oRT = new Regulador(_connBuilder, _par);

            // realiza consulta StoredReguladorMT 
            _structElem._temRegulador = oRT.ConsultaStoredReguladorMT(_genAllSubstation);

            // _temRegulador
            if (_structElem._temRegulador)
            {
                oRT.GravaEmArquivo();
            }
        }

        // cria arquivo dss de segmentos de MT
        private static void CriaSegmentoMTDSS()
        {
            SegmentoMT oSegMT = new SegmentoMT(_connBuilder, _SDEE, _par, _criaDispProtecao);

            // realiza consulta StoredSegmentoMT 
            _structElem._temSegmentoMT = oSegMT.ConsultaStoredSegmentoMT(_genAllSubstation);

            // une cabeca alimentador
            oSegMT.UneSE(_genAllSubstation);

            // _temSegmentoMT
            if (_structElem._temSegmentoMT)
            {
                oSegMT.GravaEmArquivo();

                // atualiza parametros 
                _par = oSegMT.GetParam();

                // se modo criar arq coordenadas
                if (_criaArqCoordenadas)
                {
                    //
                    oSegMT.ConsultaBusCoord(_genAllSubstation);

                    //
                    oSegMT.GravaArqCoord();
                }
            }
            // se alimentador nao tem segmento MT aborta
            else
            {
                Console.Write(_par._alim + ": sem segmento MT. Abortando!\n");
            }
        }

        private static void CriaTransformadorMTMTMTBTDSS()
        {
            Trafo oTrafo = new Trafo(_connBuilder, _par, _SDEE);

            // realiza consulta StoredReguladorMT 
            _structElem._temTransformador = oTrafo.ConsultaBanco(_genAllSubstation);

            if (_structElem._temTransformador)
            {
                oTrafo.GravaEmArquivo();
            }
        }

        private static void CriaSegmentoBTDSS()
        {
            SegmentoBT oSegBT = new SegmentoBT(_connBuilder, _par);

            // realiza consulta 
            _structElem._temSegmentoBT = oSegBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temSegmentoBT)
            {
                oSegBT.GravaEmArquivo();
            }
        }

        private static void CriaRamaisDSS()
        {
            RamalBT oRamal = new RamalBT(_connBuilder, _par);

            // realiza consulta 
            _structElem._temRamal = oRamal.ConsultaBanco(_genAllSubstation);

            if (_structElem._temRamal)
            {
                oRamal.GravaEmArquivo();
            }
        }

        private static void CriaCargaMTDSS()
        {
            if (_criaTodosOsMeses)
            {
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    // seta mes corrente
                    _iMes = i;

                    // 
                    CriaCargaMTDSSPvt();
                }
            }
            else
            {
                // 
                CriaCargaMTDSSPvt();
            }
        }

        private static void CriaCargaMTDSSPvt()
        {
            CargaMT oCargaMT = new CargaMT(_connBuilder, _iMes, _ano, _numDiasFeriadoXMes,
            _somaCurvaCargaDiariaPU, _SDEE, _par, _curvasTipicasClientesMT);

            // realiza consulta 
            _structElem._temCargaMT = oCargaMT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temCargaMT)
            {
                oCargaMT.GravaEmArquivo();
            }
        }

        private static void CriaCargaBTDSS()
        {
            if (_criaTodosOsMeses)
            {
                // TODO otimizar em uma consulta
                // repete para cada mes
                for (int i = 1; i < 13; i++)
                {
                    // seta mes corrente
                    _iMes = i;

                    //
                    CriaCargaBTDSSPvt();
                }
            }
            else
            {
                //
                CriaCargaBTDSSPvt();
            }
        }

        private static void CriaCargaBTDSSPvt()
        {
            CargaBT oCargaBT = new CargaBT(_connBuilder, _iMes, _ano, _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU,
                _SDEE, _par);

            // realiza consulta 
            _structElem._temCargaBT = oCargaBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temCargaBT)
            {
                oCargaBT.GravaEmArquivo();
            }

        }

        // cria capacitorMT
        private static void CriaCapacitorMTDSS()
        {
            CapacitorMT oCap = new CapacitorMT(_connBuilder, _par);

            // realiza consulta 
            _structElem._temCapacitorMT = oCap.ConsultaBanco(_genAllSubstation);

            if (_structElem._temCapacitorMT)
            {
                oCap.GravaEmArquivo();
            }
        }

        // cria arquivos cabecalhos ou master. 3 arquivos sao criados com o seguinte modelo, onde Alim = nome do alimentador:
        // Alim.dss -> arquivo utilizado pelo projeto ExecutorOpenDSS, com informacoes iniciais do alimentador (eg: definicao do circuito)
        // AlimAnualA.dss -> arquivo p/ ser utilizado pelo usuario na GUI do OpenDSS. 
        // AlimAnulaB.dss -> arquivo utilizado pelo projeto ExecutorOpenDSS, com informacoes finais do alimentado (eg: solve).
        private static void CriaCabecalhoDSS()
        {
            CircMT oCircMT = new CircMT(_connBuilder, _par, _genAllSubstation, _criaArqCoordenadas,
                _structElem, _iMes, _coordMT);

            // realiza consulta 
            List<string> infoCabecalho = oCircMT.ConsultaStoredCircMT();

            if (infoCabecalho != null)
            {
                // grava arquivo para ser utilizado pela OpenDSS 
                ArqManip.SafeDelete(GetNomeArqCabecalho());

                ArqManip.GravaEmArquivo(infoCabecalho[0], GetNomeArqCabecalho());

                // arquivo para ser utilizado pela customizacao COM do OpenDSS
                ArqManip.SafeDelete(GetNomeArqCabecalhoCOM());

                ArqManip.GravaEmArquivo(infoCabecalho[1], GetNomeArqCabecalhoCOM());

                // arquivo para ser utilizado pela customizacao COM do OpenDSS
                ArqManip.SafeDelete(GetNomeArquivoB());

                ArqManip.GravaEmArquivo(infoCabecalho[2], GetNomeArquivoB());
            }
        }

        private static string GetNomeArquivoB()
        {
            return _par._pathAlim + _par._alim + "AnualB.dss";
        }

        private static string GetNomeArqCabecalhoCOM()
        {
            return _par._pathAlim + _par._alim + ".dss";
        }

        private static string GetNomeArqCabecalho()
        {
            return _par._pathAlim + _par._alim + "AnualA.dss";
        }

    }
}

