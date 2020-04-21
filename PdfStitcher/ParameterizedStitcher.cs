using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace PdfStitcher
{
    public class ParameterizedStitcher
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("No parameters supplied; type 'help' for the list of switches.");
                return;
            }
            if (args[0].ToLower() == "help")
            {                
                Console.WriteLine("-t=? Type of document to stitch (Arguments: pdf/inventor)");
                Console.WriteLine("-r=? Degrees of rotation to be applied to documents: must be a multiple of 90 (Default 0)");
                Console.WriteLine("-o=? Origin directory to search (Default is current directory)");
                Console.WriteLine("-c=? Search sub folders within origin directory? (Default true)");
                Console.WriteLine("-d=? Destination directory for output pdf (Default if current directory)");
                Console.WriteLine("-n=? Name of merged file (Default is 'MERGED_DATETIME')");                
                Console.WriteLine("-f={?} Files to be merged seperated by a colon ':' (Default is all documents of selected type in orgin directory)");
                Console.WriteLine("-x=? Regex string to be parsed to get files to merge. (Takes priority over a list of files)");
                Console.WriteLine("-s=? Show the PDF after creation. (Default true)");
                return;
            }
            else
            {
                StitcherBuilder stitcherBuilder = new StitcherBuilder();
                foreach (string argument in args)
                {
                    switch (argument[1])
                    {
                        case 't':
                            stitcherBuilder.Type = StitcherBuilder.GetType(argument.Substring(3));
                            break;
                        case 'r':
                            stitcherBuilder.Rotation = int.Parse(argument.Substring(3));
                            break;
                        case 'o':
                            stitcherBuilder.Origin = argument.Substring(3);
                            break;
                        case 'c':
                            stitcherBuilder.Recursive = bool.Parse(argument.Substring(3));
                            break;
                        case 'd':
                            stitcherBuilder.Destination = argument.Substring(3);
                            break;
                        case 'n':
                            stitcherBuilder.Name = argument.Substring(3);
                            break;
                        case 'f':
                            string fileStr = argument.Substring(4);
                            stitcherBuilder.Files = fileStr.Remove(fileStr.Length - 1, 1).Split(':');
                            break;
                        case 'x':
                            stitcherBuilder.Pattern = argument.Substring(3);
                            break;
                        case 's':
                            stitcherBuilder.Show = bool.Parse(argument.Substring(3));
                            break;
                        default:
                            break;
                    }
                }

                if (stitcherBuilder.IsValid)
                {
                    PdfDocument pdfDocument = stitcherBuilder.Stitch();
                    if (pdfDocument == null)
                    {
                        Console.WriteLine("Returned a null document.");
                        return;
                    }

                    pdfDocument.Save(@stitcherBuilder.Destination + "\\" + stitcherBuilder.Name + ".pdf");
                    Console.WriteLine("\r\nPDF Created");

                    if (stitcherBuilder.Show)
                    {
                        try
                        {
                            Process.Start("acrobat", @stitcherBuilder.Destination + "\\" + stitcherBuilder.Name + ".pdf");
                        }
                        catch (Exception)
                        {
                            try
                            {
                                Process.Start("AcroRd32", @stitcherBuilder.Destination + "\\" + stitcherBuilder.Name + ".pdf");
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    Process.Start(@stitcherBuilder.Destination + "\\" + stitcherBuilder.Name + ".pdf");
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                    }
                }
                else
                    Console.WriteLine("Can not complete operation; Type or Name switch is invalid.");
            }
        }
    }
}
