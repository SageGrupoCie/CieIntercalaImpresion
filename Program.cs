using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using PDFtoPrinter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CieIntercalaImpresion
{
    class Program
    {
        static string rutaFicheroOrigen;
        static string nombreFicheroOrigen;
        static string impresora;
        static string rutaFicheroDestino;
        static string nombreFicheroDestino;

        static string rutaFicDest;

        static void Main(string[] args)
        {
            /* ASIGNAMOS PARÁMETROS A VARIABLES */
            //args[0] : Ruta fichero

            

            rutaFicheroOrigen = args[0];
            //args[1] : Nombre fichero
            nombreFicheroOrigen = args[1];
            //args[2] : Impresora
            impresora = args[2];
            




            //string[] argumentos = args[1].Split(';');
            //rutaFicheroOrigen = argumentos[0];
            //nombreFicheroOrigen = argumentos[1];
            //impresora = argumentos[2];

            /*
             rutaFicheroOrigen = @"C:\GRUPOCIE\PRUEBAANDREU2.pdf";
             nombreFicheroOrigen = "PRUEBAANDREU2.pdf";
             impresora = "RICOH Aficio MP C3001 PCL 6 PRUEBAS";
            
            //impresora = "Microsoft Print to PDF";

            */


            try
            {

                nombreFicheroDestino = nombreFicheroOrigen.Replace(".Pdf", "Cie.pdf");
                rutaFicheroDestino = rutaFicheroOrigen.Substring(0, rutaFicheroOrigen.Length - nombreFicheroOrigen.Length) + "CieImpresion";
                rutaFicDest = rutaFicheroDestino;
                if (!Directory.Exists(rutaFicheroDestino))
                {
                    Directory.CreateDirectory(rutaFicheroDestino);
                }
                rutaFicheroDestino += @"\" + nombreFicheroDestino;
                new EscribirLog(nombreFicheroDestino, false);
                new EscribirLog(rutaFicheroDestino, false);
                /* CREAMOS EL FICHERO SECUNDARIO */
                crearPDFMod();
            }
            catch (Exception ex)
            {
                new EscribirLog("Error al crear el fichero: " + ex.Message, false);
                Environment.Exit(0);

            }

            try
            {
                /* IMPRIMIMOS EL FICHERO NUEVO */
                imprimirFichero();
                /* ELIMINAMOS FICHEROS ANTIGUOS */
                //File.Delete(rutaFicheroDestino);
                /*
                int diasLimite = 2;
                foreach (var item in Directory.GetDirectories(rutaFicDest))
                {
                    if (new DirectoryInfo(item).CreationTime.Add(TimeSpan.FromDays(diasLimite)) < DateTime.Now)
                    {
                        Directory.Delete(item, true);
                    }
                }
                //quitar a partir de aqui ANDREU
                */
                /*
                int diasLimite = 1;
                DirectoryInfo di = new DirectoryInfo(rutaFicDest);
                FileInfo[] files = di.GetFiles();
                foreach (FileInfo file in files)
                {
                    if (file.CreationTime.Add(TimeSpan.FromDays(diasLimite)) < DateTime.Now)
                    {
                        file.Delete();
                    }
                }
                */

            }
            catch (Exception ex)
            {
                new EscribirLog("Error al imprimir el fichero: " + ex.Message, false);

            }
            finally
            {
                //EscribirLog.eliminarFichero(5);
                Environment.Exit(0);
            }
        }

        private static void crearPDFMod()
        {
            string file = rutaFicheroOrigen;
            string range = "";
            PdfReader pdfR = new PdfReader(file);
            pdfR.SetUnethicalReading(true);
            var pdfDocumentInvoiceNumber = new iText.Kernel.Pdf.PdfDocument(pdfR);
            int numPaginas = pdfDocumentInvoiceNumber.GetNumberOfPages();
            int contPag = 1;
            if (numPaginas % 2 != 0)
            {
                while (contPag <= numPaginas)
                {
                    if (contPag != 1)
                    {
                        range = range + "," + contPag.ToString();
                    }
                    else
                    {
                        range = contPag.ToString();
                    }

                    contPag += 1;
                }
            }
            else
            {
                while (contPag <= (numPaginas / 2))
                {
                    if (range == "")
                    {
                        range = contPag.ToString() + "," + (numPaginas - (numPaginas / 2) + contPag).ToString();
                    }
                    else
                    {
                        range = range + "," + contPag.ToString() + "," + (numPaginas - (numPaginas / 2) + contPag).ToString();
                    }

                    contPag += 1;
                }
            }
            var split = new ImprovedSplitter(pdfDocumentInvoiceNumber, pageRange => new PdfWriter(rutaFicheroDestino));
            var result = split.ExtractPageRange(new PageRange(range));
            result.Close();
            pdfDocumentInvoiceNumber.Close();

            pdfR.Close();
        }

        private static void imprimirFichero()
        {
            /*
            ProcessStartInfo startInfo = new ProcessStartInfo("PDFtoPrinter_m.exe");

            startInfo.Arguments = "\"" + rutaFicheroDestino + "\" \"" + impresora + "\" /s";
            System.Diagnostics.Process.Start(startInfo);
            */

            var allowedCocurrentPrintings = 1;
            var printer = new PDFtoPrinterPrinter(allowedCocurrentPrintings);
            //printer.Print(new PrintingOptions("/s " + impresora, rutaFicheroDestino));
            printer.Print(new PrintingOptions("/s " + impresora, rutaFicheroDestino));




            //IronPdf.ChromePdfRenderer renderered = new IronPdf.ChromePdfRenderer();
            //PdfDocument pdf = renderered.RenderUrlAsPdf("jhgfhg");


            //var printer = new PDFtoPrinterPrinter();
            //printer.Print(new PrintingOptions(impresora, rutaFicheroDestino));

            //printer = new CleanupFilesPrinter(new PDFtoPrinterPrinter());
            //printer.Print(new PrintingOptions(impresora, rutaFicheroDestino));

            // Creamos instancia de la impresora
            //IPrinter printer = new Printer();

            // Imprimimos fichero
            //printer.PrintRawFile(impresora, rutaFicheroDestino, nombreFicheroDestino);

            //RawPrinterHelper.SendFileToPrinter(impresora, rutaFicheroDestino);

        }





    }
}
