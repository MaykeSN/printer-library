using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Text;

namespace PrinterLibrary
{
    public class Printer
    {
        public Printer(string zpl, string printerName)
        {
            Zpl = zpl;
            PrinterName = printerName;
        }
        public string Zpl { get; set; }
        public string PrinterName { get; set; }
        public static List<string> GetPrinters()
        {
            List<string> printers = new List<string>();

            var pr = PrinterSettings.InstalledPrinters;

            foreach (var printer in pr)
            {
                printers.Add(printer.ToString());
            }

            return printers;
        }

        #region Imports

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern IntPtr OpenPrinter(string printerName, out IntPtr phPrinter, IntPtr pDefault);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, int Level, [In] ref DOC_INFO_1 pDocInfo);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool WritePrinter(IntPtr hPrinter, byte[] pBytes, int dwCount, out int pcWritten);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool ClosePrinter(IntPtr hPrinter);
        #endregion

        private struct DOC_INFO_1
        {
            public string pDocName;
            public string pOutputFile;
            public string pDatatype;
        }

        /// <summary>
        /// Imprime de acordo com o ZPL que foi instanciado na Impressora que foi passada no construtor
        /// </summary>
        /// <returns>Retorna se a operação foi bem sucedida</returns>
        public bool Print()
        {
            IntPtr printerHandle = IntPtr.Zero;
            try
            {
                OpenPrinter(PrinterName, out printerHandle, IntPtr.Zero);
                DOC_INFO_1 docInfo = new DOC_INFO_1
                {
                    pDocName = "ZPL Document",
                    pOutputFile = null,
                    pDatatype = "RAW"
                };

                if (!StartDocPrinter(printerHandle, 1, ref docInfo))
                    throw new Exception("Não foi possível iniciar o documento.");

                if (!StartPagePrinter(printerHandle))
                    throw new Exception("Não foi possível iniciar a página.");

                byte[] zplBytes = Encoding.UTF8.GetBytes(Zpl);

                if (!WritePrinter(printerHandle, zplBytes, zplBytes.Length, out int bytesWritten))
                    throw new Exception("Não foi possível escrever na impressora.");

                if (!EndPagePrinter(printerHandle))
                    throw new Exception("Não foi possível finalizar a página.");

                if (!EndDocPrinter(printerHandle))
                    throw new Exception("Não foi possível finalizar o documento.");

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (printerHandle != IntPtr.Zero)
                    ClosePrinter(printerHandle);
            }
        }
    }
}
