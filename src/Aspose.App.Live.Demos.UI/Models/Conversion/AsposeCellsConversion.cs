using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Diagnostics;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposeCellsConversion class to convert spreadsheet file to different format
	///</Summary>
	public class AsposeCellsConversion : ApiBase
    {
        private Response ProcessTask(string fileName, string folderName, string outFileExtension, bool createZip,  bool checkNumberofPages, ActionDelegate action)
        {
           License.SetAsposeCellsLicense();
            return  Process(this.GetType().Name, fileName, folderName, outFileExtension, createZip, checkNumberofPages,  (new StackTrace()).GetFrame(5).GetMethod().Name, action);

		}
		///<Summary>
		/// ConvertXlsToXps method to convert XLS file to XPS
		///</Summary>
		
        public  Response ConvertXlsToXps(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".xps", false,   false, delegate (string inFilePath, string outPath, string zipOutFolder) 

			{
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);
                workbook.Save(outPath, Aspose.Cells.SaveFormat.XPS);
            });
        }
		///<Summary>
		/// ConvertXlsToTiff method to convert XLS file to TIFF
		///</Summary>
		
        public Response ConvertXlsToTiff(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".tiff", false,   false,  delegate (string inFilePath, string outPath, string zipOutFolder) 
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);
                workbook.Save(outPath, Aspose.Cells.SaveFormat.TIFF);
            });
        }
		///<Summary>
		/// ConvertXlsToPDF method to convert XLS file to PDF
		///</Summary>
		
        public Response ConvertXlsToPDF(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false,   false, delegate (string inFilePath, string outPath, string zipOutFolder) 
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);                
                workbook.Save(outPath, Aspose.Cells.SaveFormat.Pdf);
            });
        }
		///<Summary>
		/// ConvertXlsToHtml method to convert XLS file to HTML
		///</Summary>
		
        public Response ConvertXlsToHtml(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".html", true,  false, delegate (string inFilePath, string outPath, string zipOutFolder) 
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);
                workbook.Save(outPath, Aspose.Cells.SaveFormat.Html);
            });
        }
		///<Summary>
		/// ConvertXlsToJpg method to convert XLS file to Jpg
		///</Summary>
		
        public Response ConvertXlsToJpg(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".jpg", true, false, delegate (string inFilePath, string outPath, string zipOutFolder) 
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);

                Aspose.Cells.Rendering.ImageOrPrintOptions imgOptions = new Aspose.Cells.Rendering.ImageOrPrintOptions();
				imgOptions.ImageType = Aspose.Cells.Drawing.ImageType.Jpeg;
                imgOptions.OnePagePerSheet = true;

                foreach (var sheet in workbook.Worksheets)
                {
                    string outfileName = sheet.Name + ".jpg";
                    outPath = zipOutFolder + "/" + outfileName;

                    Aspose.Cells.Rendering.SheetRender sr = new Aspose.Cells.Rendering.SheetRender(sheet, imgOptions);
                    System.Drawing.Bitmap bitmap = sr.ToImage(0);
                    if (bitmap != null)
                    {
                        bitmap.Save(outPath);
                    }
                }
            });            
        }
		///<Summary>
		/// ConvertXlsToPng method to convert XLS file to PNG
		///</Summary>
		
        public Response ConvertXlsToPng(string fileName, string folderName, string userEmail)
        {
            return  ProcessTask(fileName, folderName, ".png", true,   false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);

                Aspose.Cells.Rendering.ImageOrPrintOptions imgOptions = new Aspose.Cells.Rendering.ImageOrPrintOptions();
				imgOptions.ImageType = Aspose.Cells.Drawing.ImageType.Jpeg;
				imgOptions.OnePagePerSheet = true;

                foreach (var sheet in workbook.Worksheets)
                {
                    string outfileName = sheet.Name + ".png";
                    outPath = zipOutFolder + "/" + outfileName;

                    Aspose.Cells.Rendering.SheetRender sr = new Aspose.Cells.Rendering.SheetRender(sheet, imgOptions);
                    System.Drawing.Bitmap bitmap = sr.ToImage(0);
                    if (bitmap != null)
                    {
                        bitmap.Save(outPath);
                    }
                }
            });
        }
		///<Summary>
		/// ConvertXlsToSvg method to convert XLS file to SVG
		///</Summary>
		public Response ConvertXlsToSvg(string fileName, string folderName)
        {                                                            
            return  ProcessTask(fileName, folderName, ".svg", true,   true, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);

                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.SaveFormat = SaveFormat.SVG;
                imgOptions.OnePagePerSheet = true;
				int sheetCount = workbook.Worksheets.Count;
                foreach (var sheet in workbook.Worksheets)
                {
                    Aspose.Cells.Rendering.SheetRender sr = new Aspose.Cells.Rendering.SheetRender(sheet, imgOptions);
					int srPageCount = sr.PageCount;
                    for (int i = 0; i < sr.PageCount; i++)
                    {
						string outfileName = "";
						if ((sheetCount > 1) || (srPageCount > 1))
						{
							outfileName = sheet.Name + "_" + (i + 1) + ".svg";
							outPath = zipOutFolder + "/" + outfileName;
						}
						else
						{
							outfileName = System.IO.Path.GetFileNameWithoutExtension(fileName) + ".svg";
							outPath = zipOutFolder + "/" + outfileName;
						}
						sr.ToImage(i, outPath);
					}
                }
            });                
        }
		///<Summary>
		/// ConvertXlsToImageFiles method to convert XLS file to image
		///</Summary>
		public Response ConvertXlsToImageFiles(string fileName, string folderName, string outputType)
        {
            if (outputType.Equals("bmp") || outputType.Equals("jpg") || outputType.Equals("png"))
            {
				Aspose.Cells.Drawing.ImageType format = Aspose.Cells.Drawing.ImageType.Bmp;
			
                if (outputType.Equals("jpg"))
                {
                    format = Aspose.Cells.Drawing.ImageType.Jpeg;
                }
                else if (outputType.Equals("png"))
                {
                    format = Aspose.Cells.Drawing.ImageType.Png;
                }

                return  ProcessTask(fileName, folderName, "." + outputType, true,   true, delegate (string inFilePath, string outPath, string zipOutFolder)
                {
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);

                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                    imgOptions.ImageType = format;
                    imgOptions.OnePagePerSheet = true;

					int worksheetCount = workbook.Worksheets.Count;
                    foreach (var sheet in workbook.Worksheets)
                    {
						string outfileName = "";
						if (worksheetCount > 1)
						{
							outfileName = sheet.Name + "." + outputType;
						}
						else
						{
							outfileName = System.IO.Path.GetFileNameWithoutExtension(fileName) + "." + outputType;
						}
                        outPath = zipOutFolder + "/" + outfileName;

                        Aspose.Cells.Rendering.SheetRender sr = new Aspose.Cells.Rendering.SheetRender(sheet, imgOptions);
                        System.Drawing.Bitmap bitmap = sr.ToImage(0);
                        if (bitmap != null)
                        {
                            bitmap.Save(outPath);
                        }
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
		/// ConvertXlsToPDF method to convert XLS file to PDF
		///</Summary>
		public Response ConvertXlsToPDF(string fileName, string folderName, string userEmail, string outputType)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {                
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);

                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

                if (outputType == "pdfa_1b")
                {
                    pdfSaveOptions.Compliance = Cells.Rendering.PdfCompliance.PdfA1b;                    
                }
                else if (outputType == "pdfa_1a")
                {
                    pdfSaveOptions.Compliance = Cells.Rendering.PdfCompliance.PdfA1a;                
                }
                else if (outputType == "pdf_15" || outputType == "pdf")
                {
                    pdfSaveOptions.Compliance = Cells.Rendering.PdfCompliance.None;                 
                }

                workbook.Save(outPath, pdfSaveOptions);

            });
        }
		///<Summary>
		/// ConvertXlsToOds method to convert XLS file to ods
		///</Summary>
		public Response ConvertXlsToOds(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".ods", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(inFilePath);
                workbook.Save(outPath, Aspose.Cells.SaveFormat.ODS);
            });
        }
		///<Summary>
		/// ConvertExcelToExcel method to convert excel file to excel
		///</Summary>
		public Response ConvertExcelToExcel(string fileName, string folderName, string outputType)
        {
            return  ProcessTask(fileName, folderName, "." + outputType, false,   false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                var format = Aspose.Cells.SaveFormat.Xlsx;
                switch (outputType)
                {
                    case "xls":
                        format = Aspose.Cells.SaveFormat.Excel97To2003;
                        break;
                    case "xlsx":
                        format = Aspose.Cells.SaveFormat.Xlsx;
                        break;
                    case "xlsm":
                        format = Aspose.Cells.SaveFormat.Xlsm;
                        break;
                    case "xlsb":
                        format = Aspose.Cells.SaveFormat.Xlsb;
                        break;
                    case "xlam":
                        format = Aspose.Cells.SaveFormat.Xlam;
                        break;
                    case "csv":
                        format = Aspose.Cells.SaveFormat.CSV;
                        break;
                    case "tabdelimited":
                        format = Aspose.Cells.SaveFormat.TabDelimited;
                        break;
                }
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(inFilePath);
                wb.Save(outPath, format);
            });
        }
		///<Summary>
		/// ConvertFile
		///</Summary>
		public  Response ConvertFile(string fileName, string folderName, string outputType)
        {
            outputType = outputType.ToLower();

            if (outputType.StartsWith("ods"))
            {
                return  ConvertXlsToOds(fileName, folderName);
            }
            else if (outputType.StartsWith("pdf"))
            {
                return  ConvertXlsToPDF(fileName, folderName, "");
            }
            else if (outputType.Equals("xps"))
            {
                return  ConvertXlsToXps(fileName, folderName, "");
            }
            else if (outputType.Equals("html"))
            {
                return  ConvertXlsToHtml(fileName, folderName, "");
            }
            else if (outputType.Equals("jpg") || outputType.Equals("png") || outputType.Equals("bmp"))
            {
                return  ConvertXlsToImageFiles(fileName, folderName, outputType);
            }
            else if (outputType.Equals("tiff"))
            {
                return  ConvertXlsToTiff(fileName, folderName, "");
            }            
            else if (outputType.Equals("svg"))
            {
                return  ConvertXlsToSvg(fileName, folderName);
            }
            else if (outputType.Equals("xls") || outputType.Equals("xlsx") || outputType.Equals("xlsm") || outputType.Equals("xlsb") || outputType.Equals("xlam")
                                       || outputType.Equals("csv") || outputType.Equals("tabdelimited"))
            {
                return  ConvertExcelToExcel(fileName, folderName, outputType);
            }

            return new Response
            {
                FileName = null,
                Status = "Output type not found",
                StatusCode = 500
            };
        }
		///<Summary>
		/// Merge method to merge excel files
		///</Summary>
		public Response Merge(string fileName1, string fileName2, string folderName)
        {
            string outFileExtension = Path.GetExtension(fileName1).ToLower();

            License.SetAsposeCellsLicense();
            var combinedDocument = string.Format("{0}_CombineTo_{1}" + outFileExtension,
                Path.GetFileNameWithoutExtension(fileName1), Path.GetFileNameWithoutExtension(fileName2));

            return  Process(this.GetType().Name, combinedDocument, folderName, outFileExtension, false, false,  "Merge",
				(inFilePath, outPath, zipOutFolder) =>
                {
                    var wb1 = new Workbook(Aspose.App.Live.Demos.UI.Config.Configuration.WorkingDirectory + folderName + "/" + fileName1);
                    var wb2 = new Workbook(Aspose.App.Live.Demos.UI.Config.Configuration.WorkingDirectory + folderName + "/" + fileName2);
                    wb1.Combine(wb2);
                    wb1.Save(outPath);
                });
        }
    }
}
