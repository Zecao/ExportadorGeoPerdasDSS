using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS
{
    class LeXLSX
    {
        //
        public static Dictionary<string, double> XLSX2Dictionary(string nomeArquivoCompleto, int coluna = 2 )
        {
            Dictionary<string, double> saida = new Dictionary<string, double>();

            // verifica existencia de arquivo
            if (File.Exists(nomeArquivoCompleto))
            {
                var file = new FileInfo(nomeArquivoCompleto);
                
                // TODO opcao de executar com o arquivo aberto
                using (var package = new ExcelPackage(file))
                {
                    ExcelWorksheet plan = package.Workbook.Worksheets.First();
                    int ultimaLinha = plan.Dimension.End.Row;
                    double valor;
                    string texto;
                    for (int i = 1; i <= ultimaLinha; i++)
                    {
                        try
                        {
                            texto = plan.Cells[i, 1].Text;
                            valor = double.Parse(plan.Cells[i, coluna].Value.ToString());
                            saida.Add(texto, valor);
                        }
                        catch { }
                    }
                }
                return saida;
            }
            else
            {
                throw new FileNotFoundException("Arquivo " + nomeArquivoCompleto + " não encontrado.");
            }
        }
    }
}
