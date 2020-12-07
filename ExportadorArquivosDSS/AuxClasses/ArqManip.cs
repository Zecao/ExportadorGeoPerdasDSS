﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS
{
    class ArqManip
    {
        //Verifica se o arquivo existe antes de deletá-lo
        public static void SafeDelete(string arquivo)
        {
            if (File.Exists(arquivo))
            {
                File.Delete(arquivo);
            }
        } 

        //Grava CONTEUDO em arquivo FID 
        public static void GravaEmArquivo(string conteudo, string fid )
        {
            try
            {
                File.AppendAllText(fid, conteudo);
            }
            catch
            {
            }
        }
    }
}
