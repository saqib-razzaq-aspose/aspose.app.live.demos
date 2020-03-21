using System.IO;
using System.Web.Http;
using System.Threading.Tasks;
using Aspose.Pdf.Devices;
using System.Diagnostics;
using System;
using System.Linq;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposePdfConversion class to convert PDF file to different formats
	///</Summary>
	public class AsposePdfConversion : ApiBase
	{

		private Response ProcessTask(string fileName, string folderName, string outFileExtension, bool createZip, string userEmail, bool checkNumberofPages, ActionDelegate action)
		{
			License.SetAsposePdfLicense();
			//return await Process("AsposePdfConversionController", fileName, folderName, outFileExtension, createZip, checkNumberofPages,
			//    "Converter", "Aspose.PDF", "Convert", action);
			return  Process(this.GetType().Name, fileName, folderName, outFileExtension, createZip, checkNumberofPages, (new StackTrace()).GetFrame(5).GetMethod().Name, action);
		}
		///<Summary>
		/// ConvertXPSToPdf to convert XPS to Pdf
		///</Summary>

        public Response ConvertXPSToPdf(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Instantiate LoadOption object using XPS load option
                Aspose.Pdf.XpsLoadOptions options = new Aspose.Pdf.XpsLoadOptions();

                // Create document object 
                Aspose.Pdf.Document document = new Aspose.Pdf.Document(inFilePath, options);

                // Save the document in PDF format.
                document.Save(outPath);
            });
        }
		///<Summary>
		/// ConvertMDToPdf to convert MD to Pdf
		///</Summary>
		
        public Response ConvertMarkownToPdf(string fileName, string folderName, string outputType)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, "", false,
                delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Pdf.MdLoadOptions options = new Aspose.Pdf.MdLoadOptions();
                Aspose.Pdf.Document document = new Aspose.Pdf.Document(inFilePath, options);
                if (outputType != "pdf")
                {                    
                    if (outputType == "pdf/a-1b")
                        document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_1B, Pdf.ConvertErrorAction.Delete);
                    if (outputType == "pdf/a-3b")
                        document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_2B, Pdf.ConvertErrorAction.Delete);
                    if (outputType == "pdf/a-2u")
                        document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_2U, Pdf.ConvertErrorAction.Delete);
                    if (outputType == "pdf/a-3u")
                        document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_3U, Pdf.ConvertErrorAction.Delete);
                }
                document.Save(outPath);
            });
        }
		///<Summary>
		/// ConvertSVGToPdf to convert SVG to Pdf
		///</Summary>
        public Response ConvertSVGToPdf(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Instantiate LoadOption object using SVG load option
                Aspose.Pdf.LoadOptions loadopt = new Aspose.Pdf.SvgLoadOptions();

                // Create document object 
                Aspose.Pdf.Document document = new Aspose.Pdf.Document(inFilePath, loadopt);

                // Save the document in PDF format.
                document.Save(outPath);
            });
        }
		///<Summary>
		/// ConvertEPUBToPdf to convert EPUB to Pdf
		///</Summary>
        public Response ConvertEPUBToPdf(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Instantiate LoadOption object using EPUB load option
                Aspose.Pdf.EpubLoadOptions epubload = new Aspose.Pdf.EpubLoadOptions();

                // Create document object 
                Aspose.Pdf.Document document = new Aspose.Pdf.Document(inFilePath, epubload);

                // Save the document in PDF format.
                document.Save(outPath);
            });
        }
		///<Summary>
		/// ConvertPdfToDoc to convert PDF to doc
		///</Summary>
        public Response ConvertPdfToDoc(string fileName, string folderName, string userEmail, string type)
        {
            return  ProcessTask(fileName, folderName, "." + type, false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Open the source PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate DocSaveOptions object
                Aspose.Pdf.DocSaveOptions saveOptions = new Aspose.Pdf.DocSaveOptions();
                type = type.ToLower();
                if (type == "docx")
                {
                    // Specify the output format as DOCX
                    saveOptions.Format = Aspose.Pdf.DocSaveOptions.DocFormat.DocX;

                }
                else if (type == "doc")
                {
                    // Specify the output format as DOC
                    saveOptions.Format = Aspose.Pdf.DocSaveOptions.DocFormat.Doc;
                }

                pdfDocument.Save(outPath, saveOptions);
            });
        }

		///<Summary>
		/// Convert Latex to PDF
		///</Summary>
		public Response ConvertLatexToPdf(string fileName, string folderName, string outputType)
		{
			return  ProcessTask(fileName, folderName, ".pdf", false, "", false,
				delegate (string inFilePath, string outPath, string zipOutFolder)
				{
					Aspose.Pdf.LatexLoadOptions options = new Aspose.Pdf.LatexLoadOptions();					
					Aspose.Pdf.Document document = new Aspose.Pdf.Document(inFilePath, options);
					if (outputType != "pdf")
					{
						if (outputType == "pdf/a-1b")
							document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_1B, Pdf.ConvertErrorAction.Delete);
						if (outputType == "pdf/a-3b")
							document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_2B, Pdf.ConvertErrorAction.Delete);
						if (outputType == "pdf/a-2u")
							document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_2U, Pdf.ConvertErrorAction.Delete);
						if (outputType == "pdf/a-3u")
							document.Convert(new MemoryStream(), Pdf.PdfFormat.PDF_A_3U, Pdf.ConvertErrorAction.Delete);
					}
					document.Save(outPath);
				});
		}

		///<Summary>
		/// ConvertPdfToXPS to convert PDF to XPS
		///</Summary>

        public Response ConvertPdfToXPS(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".xps", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Load PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate XPS Save options
                Aspose.Pdf.XpsSaveOptions saveOptions = new Aspose.Pdf.XpsSaveOptions();

                // Save the XPS document
                pdfDocument.Save(outPath, saveOptions);
            });
        }
		///<Summary>
		/// ConvertPdfToExcel to convert PDF to Excel
		///</Summary>
        public Response ConvertPdfToExcel(string fileName, string folderName, string userEmail, string type)
        {
            type = type.ToLower();
            return  ProcessTask(fileName, folderName, "." + type, false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Open the source PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate ExcelSaveOptions object
                var saveOptions = new Aspose.Pdf.ExcelSaveOptions();
                if (type == "xlsx")
                {
                    // Specify the output format as XLSX
                    saveOptions.Format = Aspose.Pdf.ExcelSaveOptions.ExcelFormat.XLSX;
                    saveOptions.ConversionEngine = Aspose.Pdf.ExcelSaveOptions.ConversionEngines.NewEngine;

				}
                else if (type == "xml")
                {
                    // Specify the output format as SpreadsheetML
                    saveOptions.Format = Aspose.Pdf.ExcelSaveOptions.ExcelFormat.XMLSpreadSheet2003;
                }

                pdfDocument.Save(outPath, saveOptions);

            });
        }
		///<Summary>
		/// ConvertPdfToSVG to convert PDF to SVG
		///</Summary>
        public Response ConvertPdfToSVG(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".svg", true, userEmail, true, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Load PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate an object of SvgSaveOptions
                Aspose.Pdf.SvgSaveOptions saveOptions = new Aspose.Pdf.SvgSaveOptions();

                // Do not compress SVG image to Zip archive
                saveOptions.CompressOutputToZipArchive = false;

                string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";

                int pageCount = 0;
                int totalPages = pdfDocument.Pages.Count;
                // Loop through all the pages
                foreach (Aspose.Pdf.Page pdfPage in pdfDocument.Pages)
                {

                    if (totalPages > 1)
                    {
                        outPath = zipOutFolder + "/" + outfileName;
                        outPath = string.Format(outPath, pageCount + 1);
                    }
                    else
                    {
                        outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
                    }

                    Aspose.Pdf.Document newDocument = new Aspose.Pdf.Document();
                    newDocument.Pages.Add(pdfPage);
                    newDocument.Save(outPath + ".svg", saveOptions);
                    pageCount++;
                }

            });

        }
		///<Summary>
		/// ConvertPdfToEpub to convert PDF to EPUB
		///</Summary>

        public Response ConvertPdfToEpub(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".epub", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Load PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate Epub Save options
                Aspose.Pdf.EpubSaveOptions options = new Aspose.Pdf.EpubSaveOptions();

                // Specify the layout for contents
                options.ContentRecognitionMode = Aspose.Pdf.EpubSaveOptions.RecognitionMode.Flow;

                // Save the output in epub file
                pdfDocument.Save(outPath, options);
            });
        }
		///<Summary>
		/// ConvertPdfToTex to convert PDF to Tex
		///</Summary>

        public Response ConvertPdfToTex(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".tex", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Load PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate LaTex save option            
                Aspose.Pdf.LaTeXSaveOptions saveOptions = new Aspose.Pdf.LaTeXSaveOptions();

                // Specify the output directory 
                saveOptions.OutDirectoryPath = Aspose.App.Live.Demos.UI.Config.Configuration.OutputDirectory;

                // Save the output in SVG files
                pdfDocument.Save(outPath, saveOptions);
            });
        }
		///<Summary>
		/// ConvertPdfToPPTX to convert PDF to PPTX
		///</Summary>
        public Response ConvertPdfToPPTX(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".pptx", false, userEmail, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Load PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate PptxSaveOptions instance
                Aspose.Pdf.PptxSaveOptions options = new Aspose.Pdf.PptxSaveOptions();

                // Save the output in PPTX format
                pdfDocument.Save(outPath, options);
            });
        }

		///<Summary>
		/// ConvertPdfToTiff method to convert PDF to TIFF
		///</Summary>
		public Response ConvertPdfToTiff(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".tiff", false, "", false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                Resolution resolution = new Resolution(300);
                TiffDevice tiffDevice = new TiffDevice(resolution);

                tiffDevice.Process(pdfDocument, outPath);
            });
        }
		///<Summary>
		/// ConvertPdfToPng method to convert PDF to PNG
		///</Summary>
		public Response ConvertPdfToPng(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".png", true, "", true, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";
                int totalPages = pdfDocument.Pages.Count;
                for (int pageCount = 1; pageCount <= totalPages; pageCount++)
                {
                    if (totalPages > 1)
                    {
                        outPath = zipOutFolder + "/" + outfileName;
                        outPath = string.Format(outPath, pageCount);
                    }
                    else
                    {
                        outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
                    }

                    Resolution resolution = new Resolution(300);
                    PngDevice pngDevice = new PngDevice(resolution);

                    pngDevice.Process(pdfDocument.Pages[pageCount], outPath + ".png");
                }
            });
        }
		///<Summary>
		/// ConvertPdfToPng method to convert PDF to GIF
		///</Summary>
		public Response ConvertPdfToGif(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, ".gif", true, "", true, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

				string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";
				int totalPages = pdfDocument.Pages.Count;
				for (int pageCount = 1; pageCount <= totalPages; pageCount++)
				{
					if (totalPages > 1)
					{
						outPath = zipOutFolder + "/" + outfileName;
						outPath = string.Format(outPath, pageCount);
					}
					else
					{
						outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
					}

					Resolution resolution = new Resolution(300);
					GifDevice gifDevice = new GifDevice(resolution);

					gifDevice.Process(pdfDocument.Pages[pageCount], outPath + ".gif");
				}
			});
		}
		///<Summary>
		/// ConvertPdfToBmp method to convert PDF to BMP
		///</Summary>
		public Response ConvertPdfToBmp(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".bmp", true, "", true, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";
                int totalPages = pdfDocument.Pages.Count;
                for (int pageCount = 1; pageCount <= totalPages; pageCount++)
                {
                    if (totalPages > 1)
                    {
                        outPath = zipOutFolder + "/" + outfileName;
                        outPath = string.Format(outPath, pageCount);
                    }
                    else
                    {
                        outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
                    }

                    Resolution resolution = new Resolution(300);
                    BmpDevice bmpDevice = new BmpDevice(resolution);
                    bmpDevice.Process(pdfDocument.Pages[pageCount], outPath + ".bmp");
                }
            });
        }
		///<Summary>
		/// ConvertPdfToJpg method to convert PDF to JPG
		///</Summary>
		public Response ConvertPdfToJpg(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".Jpg", true, "", true, delegate (string inFilePath, string outPath, string zipOutFolder)
             {
                 Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                 string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";
                 int totalPages = pdfDocument.Pages.Count;

                 for (int pageCount = 1; pageCount <= totalPages; pageCount++)
                 {
                     if (totalPages > 1)
                     {
                         outPath = zipOutFolder + "/" + outfileName;
                         outPath = string.Format(outPath, pageCount);
                     }
                     else
                     {
                         outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
                     }

                     Resolution resolution = new Resolution(300);
                     JpegDevice jpgDevice = new JpegDevice(resolution);
                     jpgDevice.Process(pdfDocument.Pages[pageCount], outPath + ".Jpg");
                 }
             });
        }
		///<Summary>
		/// ConvertPdfToEmf method to convert PDF to Emf
		///</Summary>
		public Response ConvertPdfToEmf(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".emf", true, "", true, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";
                int totalPages = pdfDocument.Pages.Count;

                for (int pageCount = 1; pageCount <= totalPages; pageCount++)
                {
                    if (totalPages > 1)
                    {
                        outPath = zipOutFolder + "/" + outfileName;
                        outPath = string.Format(outPath, pageCount);
                    }
                    else
                    {
                        outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
                    }

                    Resolution resolution = new Resolution(300);
                    var emfDevice = new EmfDevice(resolution);
                    emfDevice.Process(pdfDocument.Pages[pageCount], outPath + ".emf");
                }
            });
        }
		///<Summary>
		/// ConvertPdfToHtml method to convert PDF to Html
		///</Summary>
		public Response ConvertPdfToHtml(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".html", true, "", false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                // Load PDF document
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inFilePath);

                // Instantiate HTMLSaveOptions instance
                Aspose.Pdf.HtmlSaveOptions options = new Aspose.Pdf.HtmlSaveOptions
                {
                    DocumentType = Pdf.HtmlDocumentType.Html5,
                    RasterImagesSavingMode = Pdf.HtmlSaveOptions.RasterImagesSavingModes.AsPngImagesEmbeddedIntoSvg
                };
                // Save the output in HTML format
                pdfDocument.Save(outPath, options);
            });
        }
		///<Summary>
		/// ConvertFile method 
		///</Summary>
		public Response ConvertFile(string fileName, string folderName, string outputType)
        {
            outputType = outputType.ToLower();

            if (outputType.Equals("pptx"))
            {
                return  ConvertPdfToPPTX(fileName, folderName, "");
            }
            else if (outputType.Equals("xlsx"))
            {
                return  ConvertPdfToExcel(fileName, folderName, "", "xlsx");
            }
            else if (outputType.Equals("xml"))
            {
                return  ConvertPdfToExcel(fileName, folderName, "", "xml");
            }
            else if (outputType.StartsWith("doc"))
            {
                return  ConvertPdfToDoc(fileName, folderName, "", outputType);
            }
            else if (outputType.Equals("xps"))
            {
                return  ConvertPdfToXPS(fileName, folderName, "");
            }
            else if (outputType.Equals("epub"))
            {
                return  ConvertPdfToEpub(fileName, folderName, "");
            }
            else if (outputType.Equals("tex"))
            {
                return  ConvertPdfToTex(fileName, folderName, "");
            }
            else if (outputType.Equals("svg"))
            {
                return  ConvertPdfToSVG(fileName, folderName, "");
            }
            else if (outputType.Equals("tiff"))
            {
                return  ConvertPdfToTiff(fileName, folderName);
            }
            else if (outputType.Equals("png"))
            {
                return  ConvertPdfToPng(fileName, folderName);
            }
			else if (outputType.Equals("gif"))
			{
				return ConvertPdfToGif(fileName, folderName);
			}
			else if (outputType.Equals("bmp"))
            {
                return  ConvertPdfToBmp(fileName, folderName);
            }
            else if (outputType.Equals("jpg"))
            {
                return  ConvertPdfToJpg(fileName, folderName);
            }
            else if (outputType.Equals("emf"))
            {
                return  ConvertPdfToEmf(fileName, folderName);
            }
            else if (outputType.Equals("html"))
            {
                return  ConvertPdfToHtml(fileName, folderName);
            }

            return new Response
            {
                FileName = null,
                Status = "Output type not found",
                StatusCode = 500
            };
        }
		
	}
}
