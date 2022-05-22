﻿using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportadorGeoPerdasDSS
{
    class AuxFunc
    {
        public static string GetNumFases(string codFase)
        {
            string ret;

            switch (codFase)
            {
                case "ABC":
                    ret = "3";
                    break;
                case "AB":
                    ret = "2";
                    break;
                case "CA":
                    ret = "2";
                    break;
                case "BC":
                    ret = "2";
                    break;
                case "A":
                    ret = "1";
                    break;
                case "B":
                    ret = "1";
                    break;
                case "C":
                    ret = "1";
                    break;
                case "ABCN":
                    ret = "3";
                    break;
                case "ABN":
                    ret = "2";
                    break;
                case "CAN":
                    ret = "2";
                    break;
                case "BCN":
                    ret = "2";
                    break;
                case "AN":
                    ret = "1";
                    break;
                case "BN":
                    ret = "1";
                    break;
                case "CN":
                    ret = "1";
                    break;
                default:
                    ret = "3";
                    break;
            }
            return ret;
        }

        // retorna string de fases padrao OpenDSS de acordo com a fase 
        public static string GetFasesDSS(string codFase)
        {
            string ret;

            switch (codFase)
            {
                case "ABC":
                    ret = ".1.2.3";
                    break;
                case "AB":
                    ret = ".1.2";
                    break;
                case "CA":
                    ret = ".3.1";
                    break;
                case "AC":
                    ret = ".1.3";
                    break;
                case "BC":
                    ret = ".2.3";
                    break;
                case "A":
                    ret = ".1";
                    break;
                case "B":
                    ret = ".2";
                    break;
                case "C":
                    ret = ".3";
                    break;
                case "ABCN":
                    ret = ".1.2.3.0";
                    break;
                case "ABN":
                    ret = ".1.2.0";
                    break;
                case "CAN":
                    ret = ".1.3.0";
                    break;
                case "BCN":
                    ret = ".2.3.0";
                    break;
                case "AN":
                    ret = ".1.0";
                    break;
                case "BN":
                    ret = ".2.0";
                    break;
                case "CN":
                    ret = ".3.0";
                    break;
                default:
                    ret = ".1.2.3";
                    break;
            }
            return ret;
        }

        // converte tensao fase-fase para fase-neutro
        public static string GetTensaoFN(string tensaoFF)
        {
            string ret = "7.967";

            //condicao de retorno
            if (tensaoFF.Equals(""))
            {
                return ret;
            }

            double tensaoFN = double.Parse(tensaoFF);

            //verifica se tensao eh igual a 34.5 ou 22.0kV
            if (tensaoFN.Equals(34.5))
            {
                ret = "19.92";  
            } 
            else if (tensaoFN.Equals(22.0))
            {
                ret = "12.70";   
            }
            return ret;
        }

        // transforma consumo em demanda de acordo com o mes e curva de carga
        public static string CalcDemanda(string consumoMes, int iMes, string ano, string curva, List<List<int>> numDiasFeriadoXMes, Dictionary<string,double> somaCurvaCargaDiariaPU)
        {
            //TODO valores default
            double somaConsumoMensalPU_DU = 22;
            double somaConsumoMensalPU_SA = 4;
            double somaConsumoMensalPU_DO = 4;

            // carrega arquivo soma consumo mensal
            if (somaCurvaCargaDiariaPU.ContainsKey(curva + "DU"))
            {
                somaConsumoMensalPU_DU = somaCurvaCargaDiariaPU[curva + "DU"];
            }
            if (somaCurvaCargaDiariaPU.ContainsKey(curva + "SA"))
            {
                somaConsumoMensalPU_SA = somaCurvaCargaDiariaPU[curva + "SA"];
            }
            if (somaCurvaCargaDiariaPU.ContainsKey(curva + "DO"))
            {
                somaConsumoMensalPU_DO = somaCurvaCargaDiariaPU[curva + "DO"];
            }

            int iAno = int.Parse(ano);

            // obtem numero de dias 
            Dictionary<string, int> numDiasMes = GetNumTipoDiasMes(iMes, iAno, numDiasFeriadoXMes);

            // calcula fator"K" de conversao entre Consumo e Demanda
            double somaConsumoMensalPU = somaConsumoMensalPU_DU * numDiasMes["DU"] + somaConsumoMensalPU_SA * numDiasMes["SA"] + somaConsumoMensalPU_DO * numDiasMes["DO"];

            // conusmo mes to double
            // OBS: multiplicacao por 1000 eh para converter de MWh/Mes do GeoPerdas para kWh do OpenDSS.
            double dConsMes = double.Parse(consumoMes) * 1000;

            //Pega a potência ativa
            double demanda = dConsMes / somaConsumoMensalPU;

            // OBS: divide demanda por 2, uma vez que o modelo atual aloca 2 cargas para cada consumidor
            demanda = demanda / 2;

            // retorna demanda
            return demanda.ToString("0.#####");
        }

        // transforma consumo em demanda de acordo com o mes e 
        public static string CalcDemandaPorFatorKdiario(string consumoMes, int iMes, string ano, string fatorKdiario)
        {
            int iAno = int.Parse(ano);

            //Pega o número de dias do mês
            int numDiasMes = DateTime.DaysInMonth(iAno, iMes);

            // calcula fator"K" de conversao entre Consumo e Demanda
            double somaConsumoMensalPU = double.Parse(fatorKdiario) * numDiasMes;

            // conusmo mes to double
            // OBS: multiplicacao por 1000 eh para converter de MWh/Mes do GeoPerdas para kWh do OpenDSS.
            double dConsMes = double.Parse(consumoMes) * 1000;

            //Pega a potência ativa
            double demanda = dConsMes / somaConsumoMensalPU;

            // OBS: divide demanda por 2, uma vez que o modelo atual aloca 2 cargas para cada consumidor
            // demanda = demanda / 2;

            // retorna demanda
            return demanda.ToString("0.#####");
        }


        //Pega os feriados do ano
        public static List<List<int>> Feriados(string ano, string arqFeriados)
        {
            List<List<int>> feriados = new List<List<int>>();

            string line;
            for (int i = 0; i < 12; i++)
            {
                feriados.Add(new List<int>());
            }
            if (File.Exists(arqFeriados))
            {
                using (StreamReader file = new StreamReader(arqFeriados))
                {
                    for (int linha = 0; linha < 12; linha++)
                    {
                        line = file.ReadLine();
                        if (!line.Equals(""))
                        {
                            string[] dias = line.Split(';');
                            foreach (string dia in dias)
                            {
                                try
                                {
                                    int aux = Int32.Parse(dia.Trim());
                                    feriados[linha].Add(aux);
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            feriados.Add(new List<int>());
                        }
                    }
                }
            }
            return feriados;
        }

        //Pega o número de dias do mês, separado por tipo de dia (Dia útil, sábado e domingo e feriado)
        public static Dictionary<string, int> GetNumTipoDiasMes(int mes, int ano,  List<List<int>> numDiasFeriadoXMes)
        {
            //Pega o número de dias do mês
            int numDiasMes = DateTime.DaysInMonth(ano, mes);

            //Inicializa a contagem dos dias
            int DU = 0; //Dia útil
            int DO = 0; //Domingo e feriado
            int SA = 0; //Sábado

            //Verifica o tipo de dia para cada dia do mês
            for (int dia = 1; dia <= numDiasMes; dia++)
            {
                //Verifica se é feriado
                if (numDiasFeriadoXMes[mes - 1].Contains(dia))
                {
                    DO++;
                }
                //Caso contrário, verifica o tipo de dia
                else
                {
                    switch (new DateTime(ano, mes, dia).DayOfWeek)
                    {
                        case DayOfWeek.Saturday:
                            SA++;
                            break;
                        case DayOfWeek.Sunday:
                            DO++;
                            break;
                        default:
                            DU++;
                            break;
                    }
                }
            }

            Dictionary<string, int> tipoDia = new Dictionary<string, int>
            {
                // Prepara a variável de saída e retorna
                { "DU", DU },
                { "SA", SA },
                { "DO", DO }
            };
            return tipoDia;
        }

        internal static string IntMes2strMes(int iMes)
        {
            string ret;
            switch (iMes)
            {
                case 1:
                    ret = "Jan";
                    break;
                case 2:
                    ret = "Fev";
                    break;
                case 3:
                    ret = "Mar";
                    break;
                case 4:
                    ret = "Abr";
                    break;
                case 5:
                    ret = "Mai";
                    break;
                case 6:
                    ret = "Jun";
                    break;
                case 7:
                    ret = "Jul";
                    break;
                case 8:
                    ret = "Ago";
                    break;
                case 9:
                    ret = "Set";
                    break;
                case 10:
                    ret = "Out";
                    break;
                case 11:
                    ret = "Nov";
                    break;
                case 12:
                    ret = "Dez";
                    break;
                default:
                    ret = "Jan";
                    break;
            }
            return ret;
        }

        // retorna o consumo de acordo com o mes
        public static string GetConsumoMesCorrente(SqlDataReader rs, int iMes)
        {
            string consumoMes;

            switch (iMes)
            {
                case 1:
                    consumoMes = rs["EnerMedid01_MWh"].ToString();
                    break;
                case 2:
                    consumoMes = rs["EnerMedid02_MWh"].ToString();
                    break;
                case 3:
                    consumoMes = rs["EnerMedid03_MWh"].ToString();
                    break;
                case 4:
                    consumoMes = rs["EnerMedid04_MWh"].ToString();
                    break;
                case 5:
                    consumoMes = rs["EnerMedid05_MWh"].ToString();
                    break;
                case 6:
                    consumoMes = rs["EnerMedid06_MWh"].ToString();
                    break;
                case 7:
                    consumoMes = rs["EnerMedid07_MWh"].ToString();
                    break;
                case 8:
                    consumoMes = rs["EnerMedid08_MWh"].ToString();
                    break;
                case 9:
                    consumoMes = rs["EnerMedid09_MWh"].ToString();
                    break;
                case 10:
                    consumoMes = rs["EnerMedid10_MWh"].ToString();
                    break;
                case 11:
                    consumoMes = rs["EnerMedid11_MWh"].ToString();
                    break;
                case 12:
                    consumoMes = rs["EnerMedid12_MWh"].ToString();
                    break;
                default:
                    consumoMes = rs["EnerMedid01_MWh"].ToString();
                    break;
            }
            return consumoMes;
        }

        // converte tensao fase-fase
        // OBS: objetivo desta funcao eh preencher o valor default 13.8 caso o parametro nao seja preenchido 
        // que ocorre com a execucao da SP Principal no GeoPerdas 
        internal static string GetTensaoFF(string tensaoFF)
        {
            // tensao FF vazia (ex. trafo nao visitado)
            if (tensaoFF.Equals(""))
            {
                return "13.8";
            }
            return tensaoFF;

            /* // OLD CODE
            // nivel de tensao default
            string ret = "13.8";

            // tensao FF vazia (ex. trafo nao visitado)
            if (tensaoFF.Equals(""))
            {
                return ret;    
            }

            //necessario transformar para double
            double tensaoFFd = double.Parse(tensaoFF); 
            
            if (tensaoFFd.Equals(34.5))
                ret = "34.5";
            else if (tensaoFFd.Equals(22.0))
                ret = "22.0";
             *             
            */
        }

        internal static string GetPotPorFase(string p)
        {
            double potencia = Double.Parse(p);
            double potPorFase = potencia / 3;

            return potPorFase.ToString();
        }
    }
}
