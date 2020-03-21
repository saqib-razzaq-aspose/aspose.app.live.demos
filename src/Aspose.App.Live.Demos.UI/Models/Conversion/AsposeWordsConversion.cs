using System;
using Aspose.Words.Saving;
using System.IO;
using System.Web.Http;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposeWordsConversion class to convert words files to different formats
	///</Summary>
	public class AsposeWordsConversion : ApiBase
	{ 

    private Response ProcessTask(string fileName, string folderName, string outFileExtension, bool createZip, bool checkNumberofPages, ActionDelegate action)
		{
			License.SetAsposeWordsLicense();
			return  Process(this.GetType().Name, fileName, folderName, outFileExtension, createZip, checkNumberofPages,  (new StackTrace()).GetFrame(5).GetMethod().Name, action);
		}

		///<Summary>
		/// ConvertDocToPDF method to convert doc file to pdf format
		///</Summary>

    public Response ConvertDocToPDF(string fileName, string folderName, string type, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".pdf", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load the document from disk.
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);
				PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
				if (type == "pdfa_1b")
				{
					pdfSaveOptions.Compliance = PdfCompliance.PdfA1b;
					// Save the document in PDF/A format.
					doc.Save(outPath, pdfSaveOptions);
				}
				else if (type == "pdfa_1a")
				{
					pdfSaveOptions.Compliance = PdfCompliance.PdfA1a;
					// Save the document in PDF/A format.
					doc.Save(outPath, pdfSaveOptions);
				}
				else if (type == "pdf_15")
				{
					pdfSaveOptions.Compliance = PdfCompliance.Pdf15;
					// Save the document in PDF/A format.
					doc.Save(outPath, pdfSaveOptions);
				}
				else if (type == "pdf")
				{
					// Save the document in PDF format.
					doc.Save(outPath, pdfSaveOptions);
				}
			});
		}
		///<Summary>
		/// ConvertDocToHTML method to convert doc file to html format
		///</Summary>		
		public Response ConvertDocToHTML(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".html", true,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load the document from disk.
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);
				HtmlSaveOptions options = new HtmlSaveOptions();

				// HtmlSaveOptions.ExportRoundtripInformation property specifies
				// Whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
				// Default value is true for HTML and false for MHTML and EPUB.
				options.ExportRoundtripInformation = true;

				// Save the document in HTML format.
				doc.Save(outPath, options);
			});
		}
		///<Summary>
		/// ConvertDocToMHTML method to convert doc file to mhtml format
		///</Summary>		
		public Response ConvertDocToMHTML(string fileName, string folderName, string userEmail)
		{
			return ProcessTask(fileName, folderName, ".mhtml", true, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load the document from disk.
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);
				

				// Save the document in HTML format.
				doc.Save(outPath, Words.SaveFormat.Mhtml);
			});
		}
		///<Summary>
		/// ConvertDocToEPUB method to convert doc file to EPUB format
		///</Summary>
		public Response ConvertDocToEPUB(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".epub", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load the document from disk.
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);

				// Save the document in EPUB format.
				doc.Save(outPath);
			});
		}
		///<Summary>
		/// ConvertDocToBMP method to convert doc file to BMP format
		///</Summary>
    public Response ConvertDocToBMP(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".bmp", true, true, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load the document from disk.
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);

				ImageSaveOptions options = new ImageSaveOptions(Aspose.Words.SaveFormat.Bmp);
				options.PageCount = 1;

				string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";

				// Save each page of the document as BMP.
				for (int i = 0; i < doc.PageCount; i++)
				{
					outPath = zipOutFolder + "/" + outfileName;
					options.PageIndex = i;
					doc.Save(string.Format(outPath, i), options);
				}

			});

		}
		///<Summary>
		/// ConvertDocToEMF method to convert doc file to EMF format
		///</Summary>

    public Response ConvertDocToEMF(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".emf", true, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load the document from disk.
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);

				ImageSaveOptions options = new ImageSaveOptions(Aspose.Words.SaveFormat.Emf);
				options.PageCount = 1;

				string outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";

				// Save each page of the document as EMF.
				for (int i = 0; i < doc.PageCount; i++)
				{
					outPath = zipOutFolder + "/" + outfileName;
					options.PageIndex = i;
					doc.Save(string.Format(outPath, i + "." + "emf"), options);
				}
			});
		}
    ///<Summary>
    /// ConvertDocToSingleImageFile method to convert doc file to single image format
    ///</Summary>

    public Response ConvertDocToSingleImageFile(string fileName, string folderName, string outputType)
		{
			if (outputType.Equals("tiff") || outputType.Equals("svg"))
			{
				Aspose.Words.SaveFormat format = outputType.Equals("tiff") ? Aspose.Words.SaveFormat.Tiff : Aspose.Words.SaveFormat.Svg;

				return  ProcessTask(fileName, folderName, "." + outputType, false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
				{
					Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);
					ImageSaveOptions options = new ImageSaveOptions(format);
					doc.Save(outPath, options);
				});
			}

			return new Response
			{
				FileName = null,
				Status = "Output type not found",
				StatusCode = 500
			};
		}
    ///<Summary>
    /// ConvertDocToImageFiles method to convert doc file to image files
    ///</Summary>   
    public Response ConvertDocToImageFiles(string fileName, string folderName, string outputType)
		{
			if (outputType.Equals("bmp") || outputType.Equals("jpg") || outputType.Equals("png"))
			{
				Aspose.Words.SaveFormat format = Aspose.Words.SaveFormat.Bmp;
				if (outputType.Equals("jpg"))
				{
					format = Aspose.Words.SaveFormat.Jpeg;
				}
				else if (outputType.Equals("gif"))
				{
					format = Aspose.Words.SaveFormat.Gif;
				}
				else if (outputType.Equals("png"))
				{
					format = Aspose.Words.SaveFormat.Png;
				}

				return  ProcessTask(fileName, folderName, "." + outputType, true,  true, delegate (string inFilePath, string outPath, string zipOutFolder)
				{
					Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);

					ImageSaveOptions options = new ImageSaveOptions(format);
					options.PageCount = 1;
					string outfileName = "";
					if (doc.PageCount > 1)
					{
						outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";

						for (int i = 0; i < doc.PageCount; i++)
						{
							outPath = zipOutFolder + "/" + outfileName;
							options.PageIndex = i;

							doc.Save(string.Format(outPath, (i + 1) + "." + outputType), options);
						}
					}
					else
					{
						outfileName = Path.GetFileNameWithoutExtension(fileName);

						outPath = zipOutFolder + "/" + outfileName;
						options.PageIndex = 0;

						doc.Save(outPath + "." + outputType, options);
					}

				});
			}

			return new Response
			{
				FileName = null,
				Status = "Output type not found",
				StatusCode = 500
			};
		}
    ///<Summary>
    /// ConvertDocToDoc method to convert doc file to doc file format
    ///</Summary>

    public Response ConvertDocToDoc(string fileName, string folderName, string outputType)
		{
			return  ProcessTask(fileName, folderName, "." + outputType, false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				var format = Aspose.Words.SaveFormat.Docx;
				switch (outputType)
				{
					case "doc":
						format = Words.SaveFormat.Doc;
						break;
					case "docx":
						format = Words.SaveFormat.Docx;
						break;
					case "odt":
						format = Words.SaveFormat.Odt;
						break;
					case "ott":
						format = Words.SaveFormat.Ott;
						break;
					case "ps":
						format = Words.SaveFormat.Ps;
						break;
					case "rtf":
						format = Words.SaveFormat.Rtf;
						break;
					case "txt":
						format = Words.SaveFormat.Text;
						break;
					case "xps":
						format = Words.SaveFormat.Xps;
						break;
					case "dot":
						format = Words.SaveFormat.Dot;
						break;
					case "docm":
						format = Words.SaveFormat.Docm;
						break;
					case "dotx":
						format = Words.SaveFormat.Dotx;
						break;
					case "dotm":
						format = Words.SaveFormat.Dotm;
						break;
					case "pcl":
						format = Words.SaveFormat.Pcl;
						break;

				}
				Aspose.Words.Document doc = new Aspose.Words.Document(inFilePath);
				doc.Save(outPath, format);
			});
		}
    ///<Summary>
    /// ConvertFile
    ///</Summary>
    public Response ConvertFile(string fileName, string folderName, string outputType)
		{
			outputType = outputType.ToLower();

			if (outputType.StartsWith("pdf"))
			{
				return  ConvertDocToPDF(fileName, folderName, outputType, "");
			}
			else if (outputType.Equals("html"))
			{
				return  ConvertDocToHTML(fileName, folderName, "");
			}
			else if (outputType.Equals("mhtml"))
			{
				return ConvertDocToMHTML(fileName, folderName, "");
			}
			else if (outputType.Equals("epub"))
			{
				return  ConvertDocToEPUB(fileName, folderName, "");
			}
			else if (outputType.Equals("emf"))
			{
				return  ConvertDocToEMF(fileName, folderName, "");
			}
			else if (outputType.Equals("jpg") || outputType.Equals("png") || outputType.Equals("bmp") || outputType.Equals("gif"))
			{
				return  ConvertDocToImageFiles(fileName, folderName, outputType);
			}
			else if (outputType.Equals("tiff") || outputType.Equals("svg"))
			{
				return  ConvertDocToSingleImageFile(fileName, folderName, outputType);
			}
			else if (outputType.Equals("doc") || outputType.Equals("docx") || outputType.Equals("ps") || outputType.Equals("odt")
							 || outputType.Equals("rtf") || outputType.Equals("txt") || outputType.Equals("xps") || outputType.Equals("dot") || outputType.Equals("docm") || outputType.Equals("dotx") || outputType.Equals("dotm") || outputType.Equals("ott") || outputType.Equals("pcl"))
			{
				return  ConvertDocToDoc(fileName, folderName, outputType);
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
