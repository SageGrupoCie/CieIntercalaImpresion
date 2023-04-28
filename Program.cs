using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using PDFtoPrinter;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
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
        static string modoPDF;
        static string usuarioSage;
        static string contUsuario;

        static string rutaFicDest;
        static string rutaFicDestPDF;

        static void Main(string[] args)
        {
            /* ASIGNAMOS PARÁMETROS A VARIABLES */
            //args[0] : Ruta fichero


            /*
            rutaFicheroOrigen = args[0];
            //args[1] : Nombre fichero
            nombreFicheroOrigen = args[1];
            //args[2] : Impresora
            impresora = args[2];
            //args[3] : ModoPdf
            modoPDF = args[3];
            //args[4] : UsuarioSage
            usuarioSage = args[4];
            //args[5] : contadorUsuario
            contUsuario = args[5];
            
            */



            //string[] argumentos = args[1].Split(';');
            //rutaFicheroOrigen = argumentos[0];
            //nombreFicheroOrigen = argumentos[1];
            //impresora = argumentos[2];


            rutaFicheroOrigen = @"C:\GRUPOCIE\PRUEBAANDREU2.pdf";
             nombreFicheroOrigen = "PRUEBAANDREU2.pdf";
             //impresora = "RICOH Aficio MP C3001 PCL 6 PRUEBAS";
            
             impresora = "OneNote for Windows 10";

            modoPDF = "S";
            usuarioSage = "1";
            contUsuario = "3";




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
            if (modoPDF != "S")
            {

                /*
                var allowedCocurrentPrintings = 1;
                var printer = new PDFtoPrinterPrinter(allowedCocurrentPrintings);
                printer.Print(new PrintingOptions("/s " + impresora, rutaFicheroDestino));
                */

                // Create the printer settings for our printer
                var printerSettings = new PrinterSettings
                {
                    PrinterName = impresora,
                    Copies = 1,
                };
                /*
                // Create our page settings for the paper size selected
                var pageSettings = new PageSettings(printerSettings)
                {
                    Margins = new Margins(0, 0, 0, 0),
                };
                foreach (PaperSize paperSize in printerSettings.PaperSizes)
                {
                    //if (paperSize.PaperName == "")
                    //{
                        pageSettings.PaperSize = paperSize;
                        break;
                    //}
                }
                */
                using (var document = PdfiumViewer.PdfDocument.Load(rutaFicheroDestino))
                {
                    using (var printDocument = document.CreatePrintDocument())
                    {
                        printDocument.PrinterSettings = printerSettings;
                        //printDocument.DefaultPageSettings = pageSettings;
                        printDocument.PrintController = new StandardPrintController();
                        printDocument.Print();
                    }
                }

            }
            else 
            {
                // SI NO EXISTE EL DIRECTORIO DEL USUARIO Y EL CONTADOR, LO CREAMOS
                rutaFicDestPDF = rutaFicDest + @"\PDFs_User-" + usuarioSage + "_" + contUsuario + "_";
                if (!Directory.Exists(rutaFicDestPDF))
                {
                    Directory.CreateDirectory(rutaFicDestPDF);
                }

                //ELIMINAMOS CARPETAS ANTIGUAS SI EXISTIERAN
                string[] dirs = Directory.GetDirectories(rutaFicDest);

                foreach (string dir in dirs)
                {
                    if ((dir.Contains(@"\PDFs_User-" + usuarioSage + "_")) && !(dir.Contains(@"\PDFs_User-" + usuarioSage + "_" + contUsuario + "_")))
                    {
                        Directory.Delete(dir, true);
                    }
                }
                // copiamos el fichero a la carpeta nueva
                File.Copy(rutaFicheroDestino, rutaFicDestPDF+@"\"+ nombreFicheroDestino);
                

            }


        }





    }
}
