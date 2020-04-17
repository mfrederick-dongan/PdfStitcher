using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using PdfSharp.Pdf;
using System.IO;
using Inventor;
using System.Text.RegularExpressions;

namespace PdfStitcher
{
    class PdfStitcher
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Commands: -Help (h/H/help/Help), -Inventor (i/I/inventor/Inventor), -Pdf (p/P/pdf/Pdf), -Dir (d/D/dir/Dir, -Pdf rotated (p-r/P-R/pdf-R/PDF-R)");
                return;
            }                
            PdfDocument pdfDocument;
            switch (args[0])
            {               
                case "-i":
                case "-I":
                case "-inventor":
                case "-Inventor":
                    string desktop = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                    string path = desktop + "\\inventortemp";
                    
                    if (!Directory.Exists(desktop + "\\inventortemp"))
                    {
                        DirectoryInfo di = Directory.CreateDirectory(path);
                        di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                    }

                    pdfDocument = InventorStitchDocuments(@path);                    

                    if (pdfDocument == null)
                    {
                        Console.WriteLine("Error: returned null document.");
                        return;
                    }

                    pdfDocument.Save(path + "\\merged.pdf");
                    Console.WriteLine("PDF Created");
                    try
                    {
                        Process.Start("acrobat", @path + "\\merged.pdf");
                    }
                    catch (Exception)
                    {
                        try
                        {
                            Process.Start("AcroRd32", @path + "\\merged.pdf");
                        }
                        catch (Exception)
                        {
                            try
                            {
                                Process.Start(@path + "\\merged.pdf");
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    break;
                case "-p":
                case "-P":
                case "-pdf":
                case "-Pdf":
                    if (args.Length == 1)
                        Console.WriteLine("A base path is needed to gather pdfs.");
                    else
                    {
                        pdfDocument = PdfStitchDocuments(args[1], args.Skip(2).ToArray());
                        if (pdfDocument == null)
                        {
                            Console.WriteLine("Error: returned null document.");
                            return;
                        }                           

                        pdfDocument.Save(@args[1] + "\\merged.pdf");
                        Console.WriteLine("PDF Created");
                        try
                        {
                            Process.Start("acrobat", @args[1] + "\\merged.pdf");
                        }
                        catch (Exception)
                        {
                            try
                            {
                                Process.Start("AcroRd32", @args[1] + "\\merged.pdf");
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    Process.Start(@args[1] + "\\merged.pdf");
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                    }
                    break;
                case "-p-r":
                case "-p-R":
                case "-P-r":
                case "-P-R":
                case "-pdf-r":
                case "-pdf-R":
                case "-PDF-r":
                case "-PDF-R":
                    if (args.Length < 2)
                        Console.WriteLine("Not all parameters have a value. {Rotate Degrees, Base Path}");
                    else
                    {
                        pdfDocument = PdfStitchDocuments(args[2], args.Skip(3).ToArray(), args[1]);
                        if (pdfDocument == null)
                        {
                            Console.WriteLine("Error: returned null document.");
                            return;
                        }

                        pdfDocument.Save(@args[1] + "\\merged.pdf");
                        Console.WriteLine("PDF Created");
                        try
                        {
                            Process.Start("acrobat", @args[1] + "\\merged.pdf");
                        }
                        catch (Exception)
                        {
                            try
                            {
                                Process.Start("AcroRd32", @args[1] + "\\merged.pdf");
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    Process.Start(@args[1] + "\\merged.pdf");
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                    }
                    break;
                case "-d":
                case "-D":
                case "-dir":
                case "-Dir":
                    if (args.Length == 1)
                        Console.WriteLine("A base path is needed to gather pdfs.");
                    else
                    {
                        pdfDocument = PdfMatchFiles(args[1], args[2]);
                        if (pdfDocument == null)
                        {
                            Console.WriteLine("Error: returned null document.");
                            return;
                        }

                        pdfDocument.Save(@args[1] + "\\merged.pdf");
                        Console.WriteLine("PDF Created");
                        try
                        {
                            Process.Start("acrobat", @args[1] + "\\merged.pdf");
                        }
                        catch (Exception)
                        {
                            try
                            {
                                Process.Start("AcroRd32", @args[1] + "\\merged.pdf");
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    Process.Start(@args[1] + "\\merged.pdf");
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                    }
                    break;
                case "-help":
                case "-Help":
                case "-h":
                case "-H":
                    if (args.Length > 1)
                    {
                        switch (args[1])
                        {
                            case "i":
                            case "I":
                            case "inventor":
                            case "Inventor":
                                Console.WriteLine("Arguments: none");
                                Console.WriteLine("Export a pdf containing all currently open documents in inventor.");
                                break;
                            case "p":
                            case "P":
                            case "pdf":
                            case "Pdf":
                                Console.WriteLine("Arguments: basePath [optional: {filenames}]");
                                Console.WriteLine("Export a pdf containing all the pdfs in a given path or of the supplied filenames");
                                break;
                            case "d":
                            case "D":
                            case "dir":
                            case "Dir":
                                Console.WriteLine("Arguments: basePath [optional: {regex}]");
                                Console.WriteLine("Export a pdf containing all the pdfs in a given path that match a starting pattern");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                        Console.WriteLine("Commands: Inventor, Pdf, Dir");
                    break;
                default:
                    pdfDocument = PdfStitchDocuments(args);
                    if (pdfDocument == null)
                    {
                        Console.WriteLine("Error: returned null document.");
                        return;
                    }

                    pdfDocument.Save(args[0].Substring(0, args[0].LastIndexOf('\\') + 1) + "\\merged.pdf");
                    Console.WriteLine("PDF Created");
                    try
                    {
                        Process.Start("acrobat", @args[1] + "\\merged.pdf");
                    }
                    catch (Exception)
                    {
                        try
                        {
                            Process.Start("AcroRd32", @args[1] + "\\merged.pdf");
                        }
                        catch (Exception)
                        {
                            try
                            {
                                Process.Start(@args[1] + "\\merged.pdf");
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                break;
            }
        }

        static PdfDocument InventorStitchDocuments(string toSavePath)
        {
            Process[] aInventor = Process.GetProcessesByName("Inventor");
            if (aInventor.Length == 0)
            {
                Console.WriteLine("Inventor application is not currently running.");
                return null;
            }                
            Application InvApp = (Application)Marshal.GetActiveObject("Inventor.Application");

            TranslatorAddIn PDFAddin = InvApp.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"] as Inventor.TranslatorAddIn;

            TranslationContext oContext = InvApp.TransientObjects.CreateTranslationContext();
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

            NameValueMap oOptions = InvApp.TransientObjects.CreateNameValueMap();
            DataMedium oDataMedium = InvApp.TransientObjects.CreateDataMedium();



            int i = 0;
            int j = 1;
            int count = InvApp.Documents.VisibleDocuments.Count;
            List<string> files = new List<string>();
            foreach (var document in InvApp.Documents.VisibleDocuments)
            {
                if (PDFAddin.HasSaveCopyAsOptions[document, oContext, oOptions])
                {
                    // Options for drawings...
                    oOptions.Value["All_Color_AS_Black"] = 1;
                    oOptions.Value["Remove_Line_Weights"] = 1;
                    oOptions.Value["Vector_Resolution"] = 400;
                    oOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintSheetRange;
                    oOptions.Value["Custom_Begin_Sheet"] = 2;
                    oOptions.Value["Custom_End_Sheet"] = 4;

                }

                files.Add(@toSavePath + "\\" + (i++).ToString() + ".pdf");
                oDataMedium.FileName = files.Last();

                PDFAddin.SaveCopyAs(document, oContext, oOptions, oDataMedium);
                Console.Write("\rExporting Inventor PDFs {0} of {1}", j++, count);
            }
            Console.Write("\n");

            PdfDocument pdfDocument = PdfStitchDocuments(@toSavePath, files.Select(s => s.Substring((s.LastIndexOf('\\')) + 1)).ToArray());

            return pdfDocument;
        }

        static PdfDocument PdfMatchFiles(string basePath, string pattern = null)
        {
            int count;
            string fileName;
            if (pattern == null)
                return PdfStitchDocuments(basePath);
            else
            {
                string[] files = Directory.GetFiles(@basePath + "\\", "*.pdf", SearchOption.AllDirectories);
                count = files.Count();
                List<string> matchedFiles = new List<string>();
                foreach (string s in files)
                {
                    int startIndex = @s.LastIndexOf('\\') + 1;
                    int length = @s.Length - startIndex - 4;
                    fileName = @s.Substring(startIndex, length);

                    if (Regex.IsMatch(fileName, pattern))
                        matchedFiles.Add(fileName);
                    //if (fileName.Substring(0, pattern.Length) == pattern)
                    //    matchedFiles.Add(fileName);
                }                

                return PdfStitchDocuments(basePath, matchedFiles.ToArray());
            }
        }
        static PdfDocument PdfStitchDocuments(string[] pdfs)
        {
            int i = 1;
            int count;
            PdfDocument document = new PdfDocument();
            PdfDocument importPdf;

            count = pdfs.Count();
            foreach (string s in pdfs)
            {
                using (importPdf = PdfSharp.Pdf.IO.PdfReader.Open(s, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                {
                    foreach (PdfPage page in importPdf.Pages)
                        document.AddPage(page);
                    Console.Write("\rStitching PDFs {0} of {1}", i++, count);
                }
            }

            Console.Write("\n");
            return document;
        }
        static PdfDocument PdfStitchDocuments(string basePath, string[] pdfs = null, string degrees = null)
        {
            int rotation = 0;
            int.TryParse(degrees, out rotation);
            int i = 1;
            int count;
            PdfDocument document = new PdfDocument();
            PdfDocument importPdf;

            if (pdfs == null || pdfs.Length == 0)
            {
                string[] files = Directory.GetFiles(@basePath + "\\", "*.pdf", SearchOption.AllDirectories);
                count = files.Count();
                foreach (string s in files)
                {
                    using (importPdf = PdfSharp.Pdf.IO.PdfReader.Open(s, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                    {                        
                        foreach (PdfPage page in importPdf.Pages)
                        {
                            page.Rotate = (page.Rotate + rotation) % 360;
                            document.AddPage(page);
                        }
                        Console.Write("\rStitching PDFs {0} of {1}", i++, count);
                    }
                }
            }
            else
            {
                count = pdfs.Count();
                foreach (string s in pdfs)
                {
                    using (importPdf = PdfSharp.Pdf.IO.PdfReader.Open(@basePath + "\\" + (s.Substring(s.Length - 4) == ".pdf" ? s : s + ".pdf"), PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                    {
                        foreach (PdfPage page in importPdf.Pages)
                        {
                            page.Rotate = (page.Rotate + rotation) % 360;
                            document.AddPage(page);
                        }
                        Console.Write("\rStitching PDFs {0} of {1}", i++, count);
                    }
                }
            }
            Console.Write("\n");
            return document;
        }
    }
}
