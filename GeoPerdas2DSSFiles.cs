using ExportadorGeoPerdasDSS;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExportadorArqDSS
{
    class GeoPerdas2DSSFiles
    {
        // codbase
        public static string _codBase = "2021124950"; //"2021124950"; //"2021129999"; //  "2021124950" //"2021124950" "2021124950" //"2020015050"; // "4950"; //"2020014950";
        public static string _schema = "geo."; //"geo."; //"";

        // mes e ano para a geracao dos arquivos de carga BT e MT
        public static int _iMes = 12;
        public static string _ano = "2021"; // 2020

        public static bool _criaTodosOsMeses = false;  // flag p/ criar todos os meses de carga MT BT e geradores
        public static bool _criaArqCoordenadas = true; // flag p/ criar arq coordenadas
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
        //public static readonly string _path = @"F:\DropboxZecao\Dropbox2018\Dropbox\01_4950_P\";
        //public static readonly string _path = @"F:\DropboxZecao\Dropbox2018\Dropbox\01_4950_RECONF\";
        //public static readonly string _path = @"F:\DropboxZecao\Dropbox2018\Dropbox\Dropbox\0doutorado\0soft\0alimCemig\1CemigDFeeders\";
        //public static readonly string _path = @"F:\DropboxZecao\Dropbox2018\Dropbox\Dropbox\0doutorado\0soft\0alimCemig\1CemigDSAIDI\";
        //public static readonly string _path = @"\0_alimTese\t2\";

        //CEMIG
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_P\";
        public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\49502\";        
        //public static readonly string _path = @"D:\Users\c055896.NETCEMIG\Desktop\01_4950\";
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_TS\";
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_atipicos_Leandro\";
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_atipicos_UsinasLM1\";        
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_atipicos_LeandroCopia2\";
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_atipicos_SemUsinas\";
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_BDGDDaimon\";
        //public static readonly string _path = @"I:\SA\GRMP\PERDAS-TECNICAS\0perdasTecnicasOpenDSS\2021\01_4950_reconf\";

        // TODO FIX this hard coded 
        // sub diretorio recursos permanentes
        public static string _permRes = "0PermRes\\";

        // servidor SGBD
        public static string _banco = "GEO_SIGR_DBPERTEC01"; //"GEO_SIGR_DBPERTEC01"; //"GEOPERDAS_2S2021"; // GEO_SIGR_DBPERTEC01 "GEO_SIGR_PERDAS_2S2021"; //GEOPERDAS_2S2021 //GEOPERDAS_2020 ""; //"GEOPERDAS_2019";  

        // banco
        public static string _dataSource = @"PWNBS-PERTEC01\PTEC"; //@"sa-corp-sql0";

        //Modelo PADRAO (GeoPerdas ANEEL)
        //OBS: Capacitor pode ser colocado na hora da execucao
        //private static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: false, modeloCarga: "ANEEL",
        //    reatanciaTrafos: false);

        // sim3 modelo de carga PCONST
        private readonly static ModeloSDEE _SDEE = new ModeloSDEE(usarCondutoresSeqZero: false, utilizarCurvaDeCargaClienteMTIndividual: false, incluirCapacitoresMT: false, modeloCarga: "PCONST",
           reatanciaTrafos: true);

        // FIM membros publicos -> parametros configuraveis usuario

        //membros privados
        private static string _alim;

        //Param(string path, string permRes, string codBase, string pathAlim, string conjAlim, string trEM)

        public static Param _par = new Param(_path, _permRes, _codBase,"","","",_modelo4condutores,_schema);

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

        static void Main()
        {
            // parametros banco de dados
            _connBuilder = new SqlConnectionStringBuilder();
            _connBuilder.DataSource = _dataSource;
            _connBuilder.InitialCatalog = _banco;
            _connBuilder.IntegratedSecurity = true;

            // variaveis auxiliares
            CarregaVariaveisAux();

            // se modo reconfiguracao
            if (_genAllSubstation)
            {
                CreateSubstationDSSFiles();
            }
            else
            {
                CreateFeederDSSFiles();
            }

            Console.Write("Fim!");
            Console.ReadKey();
        }

        // Generates entire substation
        private static void CreateSubstationDSSFiles()
        {
            // populates lstAlim with all feeders from txt file "Alimentadores.m"

            List<string> lstAlim = CemigFeeders.GetAllFeedersFromTxtFile(GetNomeArqLstAlimentadores());

            // para cada SE da lista
            foreach (string alim in lstAlim)
            {
                // get substation
                string substation = System.Text.RegularExpressions.Regex.Replace(alim, @"[\d-]", string.Empty);

                //
                string lstAlimSE = CemigFeeders.GetAllFeedersFromSubstationString(substation, _connBuilder, _par);

                if (lstAlimSE == null)
                {
                    Console.Write(substation + " não encontrado!\n");
                    continue;
                }
                //
                CriaArquivosDSS(alim, lstAlimSE);
            }
        }

        // Generates entire substation
        private static void CreateFeederDSSFiles()
        {
            // lista de alimentadores
            List<string> lstAlim = CemigFeeders.GetAllFeedersFromTxtFile(GetNomeArqLstAlimentadores());

            // para cada alimentador da lista
            foreach (string alim in lstAlim)
            {
                CriaArquivosDSS(alim);
            }
        }        

        private static string GetNomeArqLstAlimentadores()
        {
            return _path + _permRes + _arqLstAlimentadores;
        }

        // cria arquivos DSS
        private static void CriaArquivosDSS(string alim, string lstAlim ="")
        {
            // preenche variavel da classe
            _alim = alim;            
            _par._pathAlim =  _path + alim + "\\";
            _par._conjAlim = lstAlim;

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
            if ( _SDEE._incluirCapacitoresMT )
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
            
            // arquivo cabecalho
            CriaCabecalhoDSS();
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

            // 
            _par._codBase = _codBase;
        }

        private static string GetNomeArqCurvaCargaCliMT()
        {
            return _path + _par._permRes + _arqCurvaCargaCliMT;
        }
                
        private static string GetNomeArqConsumoMensalPU()
        {
            return _path + _par._permRes +_arqConsumoMensalPU;
        }

        // nome arquivo feriado
        private static string GetNomeArqFeriado()
        {
            return _path + _par._permRes + _feriado + _ano + ".txt";
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
            GeradorMT oGerMT = new GeradorMT(_alim, _connBuilder, _par, _iMes);

            // realiza consulta  
            _structElem._temGeradorMT = oGerMT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temGeradorMT)
            {
                oGerMT.GravaEmArquivo();
            } 
        }

        private static void CriaChaveMT()
        {
            ChaveMT oChaveMT = new ChaveMT(_alim, _connBuilder, _par, _criaDispProtecao);

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
            Regulador oRT = new Regulador(_alim, _connBuilder, _par);

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
            SegmentoMT oSegMT = new SegmentoMT(_alim, _connBuilder, _par, _criaDispProtecao );

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
                Console.Write(_alim + ": sem segmento MT. Abortando!\n");
            }
        }

        private static void CriaTransformadorMTMTMTBTDSS()
        {
            Trafo oTrafo = new Trafo(_alim, _connBuilder, _par, _SDEE);

            // realiza consulta StoredReguladorMT 
            _structElem._temTransformador = oTrafo.ConsultaBanco(_genAllSubstation);

            if (_structElem._temTransformador)
            {
                oTrafo.GravaEmArquivo();
            }
        }

        private static void CriaSegmentoBTDSS()
        {
            SegmentoBT oSegBT = new SegmentoBT(_alim, _connBuilder, _par);

            // realiza consulta 
            _structElem._temSegmentoBT = oSegBT.ConsultaBanco(_genAllSubstation);

            if (_structElem._temSegmentoBT)
            {
                oSegBT.GravaEmArquivo();
            }
        }

        private static void CriaRamaisDSS()
        {
            RamalBT oRamal = new RamalBT(_alim, _connBuilder, _par);

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
            CargaMT oCargaMT = new CargaMT(_alim, _connBuilder, _iMes, _ano, _numDiasFeriadoXMes,
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
            CargaBT oCargaBT = new CargaBT(_alim, _connBuilder, _iMes, _ano, _numDiasFeriadoXMes, _somaCurvaCargaDiariaPU,
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
            CapacitorMT oCap = new CapacitorMT(_alim, _connBuilder, _par);

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
            CircMT oCircMT = new CircMT(_alim, _connBuilder, _par, _genAllSubstation, _criaArqCoordenadas,
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
            return _par._pathAlim + _alim + "AnualB.dss";
        }

        private static string GetNomeArqCabecalhoCOM()
        {
            return _par._pathAlim + _alim + ".dss";
        }

        private static string GetNomeArqCabecalho()
        {
            return _par._pathAlim + _alim + "AnualA.dss";
        }

    }
}

