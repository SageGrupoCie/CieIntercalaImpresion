using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CieIntercalaImpresion
{
    class EscribirLog
    {
        public string mensajeLog { get; set; }
        public Boolean mostrarConsola { get; set; }

        //Constructor si se pasa el mensaje por parámetro en la creación de la clase
        public EscribirLog(string mensajeEnviar, Boolean mostrarConsola)
        {
            mensajeLog = mensajeEnviar;
            if (mostrarConsola)
                monstrarMensajeConsola();
            escribirLineaFichero();
        }

        //Constructor si se pasa el mensaje por setter tras la creación de la clase
        public EscribirLog()
        {
            if (mostrarConsola)
                monstrarMensajeConsola();
            escribirLineaFichero();
        }

        public void monstrarMensajeConsola()
        {
            //Quitar posibles saltos de línea del mensaje
            mensajeLog = mensajeLog.Replace(Environment.NewLine, " | ");
            mensajeLog = mensajeLog.Replace("\r\n", " | ").Replace("\n", " | ").Replace("\r", " | ");
            Console.WriteLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + " " + mensajeLog);
        }

        public static void eliminarFichero(int maxDias)
        {
            
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory +
                    "estado.log"))
            {
                if (new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory +
                    "estado.log").CreationTime.Add(TimeSpan.FromDays(maxDias)) < DateTime.Now)
                {
                    Directory.Delete(AppDomain.CurrentDomain.BaseDirectory +
                        "estado.log", true);
                }
            }
            
        }

        //Escribe el mensaje de la propiedad mensajeLog en un fichero en la carpeta del ejecutable
        public void escribirLineaFichero()
        {
            try
            {
                FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory +
                    "estado.log", FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter m_streamWriter = new StreamWriter(fs);
                m_streamWriter.BaseStream.Seek(0, SeekOrigin.End);
                //Quitar posibles saltos de línea del mensaje
                mensajeLog = mensajeLog.Replace(Environment.NewLine, " | ");
                mensajeLog = mensajeLog.Replace("\r\n", " | ").Replace("\n", " | ").Replace("\r", " | ");
                m_streamWriter.WriteLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + " " + mensajeLog);
                m_streamWriter.Flush();
                m_streamWriter.Close();
            }
            catch
            {
                //Silenciosa
            }
        }
    }
}
