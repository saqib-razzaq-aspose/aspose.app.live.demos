using System.IO;
using System.Diagnostics;
using Aspose.Page.EPS;
using Aspose.Page.EPS.Device;
using Aspose.Page.XPS;
using Aspose.Page.XPS.Presentation.Image;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposePageConversion class to convert page file to other format
	///</Summary>
	public class AsposePageConversion : ApiBase
    {
        private Response ProcessTask(string fileName, string folderName, string outFileExtension, bool createZip, bool checkNumberofPages, ActionDelegate action)
        {
            License.SetAsposePageLicense();
            return  Process(this.GetType().Name, fileName, folderName, outFileExtension, createZip, checkNumberofPages,  (new StackTrace()).GetFrame(5).GetMethod().Name, action);
        }
		///<Summary>
		/// ConvertEpsToPdf method to convert eps file to pdf
		///</Summary>
		public Response ConvertEpsToPdf(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
           {
               FileStream psStream = new FileStream(inFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
               FileStream pdfStream = new FileStream(outPath, System.IO.FileMode.Create, System.IO.FileAccess.Write);

               PsDocument document = new PsDocument(psStream);
               PdfSaveOptions options = new PdfSaveOptions(true);
               PdfDevice device = new PdfDevice(pdfStream);

               try
               {
                   document.Save(device, options);
               }
               finally
               {
                   psStream.Close();
                   pdfStream.Close();
               }

           });
        }
		///<Summary>
		/// ConvertXpsToPdf method to convert Xps file to pdf
		///</Summary>
		public Response ConvertXpsToPdf(string fileName, string folderName)
        {
            return  ProcessTask(fileName, folderName, ".pdf", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
            {
                using (FileStream pdfStream = new FileStream(outPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                using (FileStream xpsStream = new FileStream(inFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {

                    XpsDocument document = new XpsDocument(xpsStream, new XpsLoadOptions());
                    Aspose.Page.XPS.Presentation.Pdf.PdfDevice device = new Aspose.Page.XPS.Presentation.Pdf.PdfDevice(pdfStream);                    
                    document.Save(device, new Aspose.Page.XPS.Presentation.Pdf.PdfSaveOptions());
                }

            });
        }

        private Response ConvertEpsToImage(string fileName, string folderName, string outputType)
        {
            return  ProcessTask(fileName, folderName, "." + outputType, true, false,

                delegate (string inFilePath, string outPath, string zipOutFolder)
                {
                    FileStream psStream = new FileStream(inFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

                    System.Drawing.Imaging.ImageFormat imageFormat = System.Drawing.Imaging.ImageFormat.Tiff;

                    switch (outputType)
                    {
                        case "bmp":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
                            break;
                        case "png":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Png;
                            break;
                        case "jpg":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                            break;
                    }

                    PsDocument document = new PsDocument(psStream);
                    Aspose.Page.EPS.Device.ImageSaveOptions options = new Aspose.Page.EPS.Device.ImageSaveOptions(true);
                    Aspose.Page.EPS.Device.ImageDevice device = new Aspose.Page.EPS.Device.ImageDevice();
                    //Aspose.Page.EPS.Device.ImageDevice device = new Aspose.Page.EPS.Device.ImageDevice(new System.Drawing.Size(595, 842), imageFormat);

                    try
                    {
                        document.Save(device, options);
                    }
                    finally
                    {
                        psStream.Close();
                    }

                    byte[][] imagesBytes = device.ImagesBytes;

                    int i = 0;
                    foreach (byte[] imageBytes in imagesBytes)
                    {
                        string imagePath = Path.GetFileNameWithoutExtension(fileName) + "_" + (i + 1) + "." + outputType;
                        outPath = zipOutFolder + "/" + imagePath;

                        using (FileStream fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                        {
                            fs.Write(imageBytes, 0, imageBytes.Length);
                        }
                        i++;
                    }
                });
        }


        private Response ConvertXpsToImage(string fileName, string folderName, string outputType)
        {
            return  ProcessTask(fileName, folderName, "." + outputType, true, false,

                delegate (string inFilePath, string outPath, string zipOutFolder)
                {
                    using (Stream xpsStream = System.IO.File.Open(inFilePath, FileMode.Open, FileAccess.Read))
                    {

                        XpsDocument document = new XpsDocument(xpsStream, new XpsLoadOptions());
                        Aspose.Page.XPS.Presentation.Image.ImageSaveOptions options = new BmpSaveOptions();

                        switch (outputType)
                        {
                            case "tiff":
                                options = new TiffSaveOptions();
                                break;
                            case "png":
                                options = new PngSaveOptions();
                                break;
                            case "jpg":
                                options = new JpegSaveOptions();
                                break;
                        }

                        Page.XPS.Presentation.Image.ImageDevice device = new Page.XPS.Presentation.Image.ImageDevice();

                        document.Save(device, options);

                        for (int i = 0; i < device.Result.Length; i++)
                        {
                            for (int j = 0; j < device.Result[i].Length; j++)
                            {
                                string imagePath = Path.GetFileNameWithoutExtension(fileName) + "_" + (j + 1) + "." + outputType;
                                outPath = zipOutFolder + "/" + imagePath;

                                using (Stream imageStream = System.IO.File.Open(outPath, FileMode.Create, FileAccess.Write))
                                {
                                    imageStream.Write(device.Result[i][j], 0, device.Result[i][j].Length);
                                }
                            }
                        }
                    }

                });
        }
		///<Summary>
		/// ConvertFile
		///</Summary>
		public Response ConvertFile(string fileName, string folderName, string outputType)
        {
            fileName = fileName.ToLower();
            outputType = outputType.ToLower();

            if (fileName.EndsWith(".xps"))
            {
                if (outputType.StartsWith("pdf"))
                {
                    return  ConvertXpsToPdf(fileName, folderName);
                }
                else if (outputType.Equals("tiff") || outputType.Equals("png") || outputType.Equals("bmp") || outputType.Equals("jpg"))
                {
                    return  ConvertXpsToImage(fileName, folderName, outputType);
                }
            }
            else
            {
                if (outputType.StartsWith("pdf"))
                {
                    return  ConvertEpsToPdf(fileName, folderName);
                }
                else if (outputType.Equals("tiff") || outputType.Equals("png") || outputType.Equals("bmp") || outputType.Equals("jpg"))
                {
                    return  ConvertEpsToImage(fileName, folderName, outputType);
                }
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
