using Aspose.Html.Rendering.Image;
using System.Threading.Tasks;
using System.Diagnostics;
using System;
using Aspose.App.Live.Demos.UI.Helpers.Html;
using Aspose.App.Live.Demos.UI.Helpers.Html.Conversion;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposeHtmlConversionController class to convert HTML files to different formats
	///</Summary>
	public class AsposeHtmlConversion : ApiBase
    {
        private Response ProcessTask(string fileName, string folderName, string outFileExtension, bool createZip,  bool checkNumberofPages, ActionDelegate action)
        {
           License.SetAsposeHtmlLicense();            			
            return  Process(this.GetType().Name, fileName, folderName, outFileExtension, createZip, checkNumberofPages,  (new StackTrace()).GetFrame(5).GetMethod().Name,  action);

		}
		///<Summary>
		/// ProcessZipArchiveFile to process zip Archive file
		///</Summary>
		public void ProcessZipArchiveFile(ref string fileName, string folderName)
        {
            //If the input file is not zip file then return;
            if(System.IO.Path.GetExtension(fileName).ToLower()!=".zip")
            {
                return;
            }
            //Extract zip file contents and prepare them for conversion.
            string destinationDirectoryName = Aspose.App.Live.Demos.UI.Config.Configuration.WorkingDirectory + folderName + "/";
            string sourceArchiveFileName = destinationDirectoryName + fileName;
            System.IO.Compression.ZipFile.ExtractToDirectory(sourceArchiveFileName, destinationDirectoryName);

            string sourceFileName="", destFileName="";

            string[] dirFiles = System.IO.Directory.GetFiles(destinationDirectoryName);
            string inExtensions = ".htm--.html--.xhtml--.mhtml--.epub--.svg--end";

            for (int i = 0; i < dirFiles.Length; i++)
            {
                sourceFileName = dirFiles[i];

                string fileExt = System.IO.Path.GetExtension(sourceFileName).ToLower();

                if (inExtensions.Contains(fileExt + "--")==true)
                {
                    destFileName = destinationDirectoryName + System.IO.Path.GetFileNameWithoutExtension(fileName) + fileExt;
                    System.IO.File.Move(sourceFileName, destFileName);
                    break;
                }
            }
            fileName = System.IO.Path.GetFileName(destFileName);
        }
		///<Summary>
		/// ConvertHtmlToPdf to convert html file to pdf
		///</Summary>
		public Response ConvertHtmlToPdf(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
				Aspose.Html.Rendering.Pdf.PdfRenderingOptions pdf_options = new Aspose.Html.Rendering.Pdf.PdfRenderingOptions();
				SourceFormat srcFormat = ExportHelper.GetSourceFormatByFileName(fileName);
				ExportHelper helper = ExportHelper.GetHelper(srcFormat, ExportFormat.PDF);
				helper.Export(inFilePath, outPath, pdf_options);   
			});
        }
		///<Summary>
		/// ConvertHtmlToXps to convert html file to xps
		///</Summary>
		public Response ConvertHtmlToXps(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".xps", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
				Aspose.Html.Rendering.Xps.XpsRenderingOptions xps_options = new Aspose.Html.Rendering.Xps.XpsRenderingOptions();
				SourceFormat srcFormat = ExportHelper.GetSourceFormatByFileName(fileName);
				ExportHelper helper = ExportHelper.GetHelper(srcFormat, ExportFormat.XPS);
				helper.Export(inFilePath, outPath, xps_options);
            });
        }
		///<Summary>
		/// ConvertHtmlToMhtml to convert html file to mhtml
		///</Summary>
		public Response ConvertHtmlToMhtml(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".mhtml", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Html.HTMLDocument document = new Aspose.Html.HTMLDocument(inFilePath);
                document.Save(outPath, Html.Saving.HTMLSaveFormat.MHTML);
            });
        }
		///<Summary>
		/// ConvertHtmlToMarkdown to convert html file to Markdown
		///</Summary>
		public Response ConvertHtmlToMarkdown(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".md", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                Aspose.Html.HTMLDocument document = new Aspose.Html.HTMLDocument(inFilePath);
                document.Save(outPath, Html.Saving.HTMLSaveFormat.Markdown);
            });
        }

		///<Summary>
		/// Convert Markdown to HTML
		///</Summary>
		public Response ConvertMarkdownToHtml(string fileName, string folderName, string outputType)
        {
			return  ProcessTask(fileName, folderName, ".html", false, false,
				delegate (string inFilePath, string outPath, string zipOutFolder)
				{
					var document = Aspose.Html.Converters.Converter.ConvertMarkdown(inFilePath);					
					document.Save(outPath);
				});
		}
		///<Summary>
		/// ConvertHtmlToTiff to convert html file to tiff
		///</Summary>
		public Response ConvertHtmlToTiff(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".tiff", false,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
				ImageRenderingOptions img_options = new ImageRenderingOptions();
				img_options.Format = ImageFormat.Tiff;
				SourceFormat srcFormat = ExportHelper.GetSourceFormatByFileName(fileName);
				ExportHelper helper = ExportHelper.GetHelper(srcFormat, ExportFormat.TIFF);
				helper.Export(inFilePath, outPath, img_options);
            });                   
        }

		///<Summary>
		/// ConvertHtmlToImages to convert html file to images
		///</Summary>
		public Response ConvertHtmlToImages(string fileName, string folderName, string outputType)
        {
            if (outputType.Equals("bmp") || outputType.Equals("jpg")
				|| outputType.Equals("png") || outputType.Equals("gif"))
            {
				ImageFormat format = ImageFormat.Bmp;
				ExportFormat expFormat = ExportFormat.BMP;

				if (outputType.Equals("jpg"))
				{
					format = ImageFormat.Jpeg;
					expFormat = ExportFormat.JPEG;
				}
				else if (outputType.Equals("png"))
				{
					format = ImageFormat.Png;
					expFormat = ExportFormat.PNG;
				}
				else if (outputType.Equals("gif"))
				{
					format = ImageFormat.Gif;
					expFormat = ExportFormat.GIF;
				}

				return  ProcessTask(fileName, folderName, "." + outputType,true,  false, delegate (string inFilePath, string outPath, string zipOutFolder)
                {
					ImageRenderingOptions img_options = new ImageRenderingOptions();
					img_options.Format = format;
					SourceFormat srcFormat = ExportHelper.GetSourceFormatByFileName(fileName);
					ExportHelper helper = ExportHelper.GetHelper(srcFormat, expFormat);
					helper.Export(inFilePath, outPath, img_options);
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
		/// ConvertFile
		///</Summary>
		public Response ConvertFile(string fileName, string folderName, string outputType)
        {
            if (System.IO.Path.GetExtension(fileName).ToLower() == ".zip")
            {
                ProcessZipArchiveFile(ref fileName, folderName);
            }

            outputType = outputType.ToLower();

            if (outputType.StartsWith("pdf"))
            {
                return  ConvertHtmlToPdf(fileName, folderName);
            }
            else if (outputType.Equals("mhtml"))
            {
                return  ConvertHtmlToMhtml(fileName, folderName);
            }
            else if (outputType.Equals("tiff"))
            {
                return  ConvertHtmlToTiff(fileName, folderName);
            }
            else if (outputType.Equals("bmp") || outputType.Equals("jpg") || outputType.Equals("png") || outputType.Equals("gif"))
            {
                return  ConvertHtmlToImages(fileName, folderName, outputType);
            }
            else if (outputType.Equals("md"))
            {
                return  ConvertHtmlToMarkdown(fileName, folderName);
            }
            else if (outputType.Equals("xps"))
            {
                return  ConvertHtmlToXps(fileName, folderName);
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
