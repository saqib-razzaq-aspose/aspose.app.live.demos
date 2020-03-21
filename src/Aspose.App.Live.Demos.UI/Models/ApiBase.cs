using System;
using System.IO;
using System.Web.Http;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Drawing;
using System.Drawing.Imaging;

namespace  Aspose.App.Live.Demos.UI.Models
{
	///<Summary>
	/// ApiControllerBase class to have base methods
	///</Summary>

	public abstract class ApiBase : ApiController
    {
		///<Summary>
		/// ActionDelegate
		///</Summary>
		protected delegate void ActionDelegate(string inFilePath, string outPath, string zipOutFolder);
		///<Summary>
		/// inFileActionDelegate
		///</Summary>
		protected delegate void inFileActionDelegate(string inFilePath);
		///<Summary>
		/// Get File extension
		///</Summary>
		protected string GetoutFileExtension(string fileName, string folderName)
        {
			string sourceFolder = Aspose.App.Live.Demos.UI.Config.Configuration.WorkingDirectory + folderName;
            fileName = sourceFolder + "\\" + fileName;
            return Path.GetExtension(fileName);
        }
		
        protected Response Process(string controllerName, string fileName, string folderName, string outFileExtension, bool createZip, bool checkNumberofPages,  string methodName, ActionDelegate action,
      bool deleteSourceFolder = true, string zipFileName = null)
        {
            
            string guid = Guid.NewGuid().ToString();
            string outFolder = "";
			string sourceFolder = Aspose.App.Live.Demos.UI.Config.Configuration.WorkingDirectory + folderName;			
            fileName = sourceFolder + "\\" + fileName;

            string fileExtension = Path.GetExtension(fileName).ToLower();
            // Check if tiff file have more than one number of pages to create zip file or not
            if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeImagingConversionController") && ( Path.GetExtension(fileName).ToLower() == ".tiff" || Path.GetExtension(fileName).ToLower() == ".tif") )
            {
                // Get the frame dimension list from the image of the file and 
                Image _image = Image.FromFile(fileName);
                // Get the globally unique identifier (GUID) 
                Guid objGuid = _image.FrameDimensionsList[0];
                // Create the frame dimension 
                FrameDimension dimension = new FrameDimension(objGuid);
                // Gets the total number of frames in the .tiff file 
                int noOfPages = _image.GetFrameCount(dimension);
                createZip = noOfPages > 1;
                _image.Dispose();
            }				

			// Check word file have more than one number of pages or not to create zip file
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeWordsConversionController"))
			{
				Aspose.Words.Document doc = new Aspose.Words.Document(fileName);
				createZip = doc.PageCount > 1;
			}
			// Check presentation file have one or more slides to create zip file
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeSlidesAPIsController"))
			{
				Aspose.Slides.Presentation presentation = new Aspose.Slides.Presentation(fileName);
				createZip = presentation.Slides.Count > 1;
			}
			// Check visio file have one or more pages to create zip file
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeDiagramConversionController"))
			{
				Aspose.Diagram.Diagram diagram = new Aspose.Diagram.Diagram(fileName);
				createZip = diagram.Pages.Count > 1;
			}
			// Check email file have one or more pages to create zip file
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeEmailConversionController"))
			{
				Aspose.Email.MailMessage msg = Aspose.Email.MailMessage.Load(fileName);

				MemoryStream msgStream = new MemoryStream();
				msg.Save(msgStream, Aspose.Email.SaveOptions.DefaultMhtml);
				Aspose.Words.Document document = new Aspose.Words.Document(msgStream);

				createZip = document.PageCount > 1;
			}
			//Check excel file have more than on workseets to create zip or not
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeCellsAPIsController") && (outFileExtension != ".svg"))
			{
				Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(fileName);
				createZip = workbook.Worksheets.Count > 1;
			}
			//Check note file have more than on pages to create zip or not
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeNoteConversionController"))
			{
				Aspose.Note.Document document = new Aspose.Note.Document(fileName);
				int count = document.GetChildNodes<Aspose.Note.Page>().Count;
				createZip = count > 1;
			}
			//Check pdf file have more than on pages to create zip or not
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposePdfConversionController"))
			{
				Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(fileName);
				createZip = pdfDocument.Pages.Count > 1;
			}
			//Check excel file have more than on workseets to create zip or not
			else if ((checkNumberofPages) && (createZip) && (controllerName == "AsposeCellsAPIsController") && (outFileExtension == ".svg"))
			{
				Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(fileName);
				if (workbook.Worksheets.Count > 1)
				{
					createZip = true;
				}
				else
				{
					Aspose.Cells.Rendering.ImageOrPrintOptions imgOptions = new Aspose.Cells.Rendering.ImageOrPrintOptions();
					imgOptions.OnePagePerSheet = true;
					Aspose.Cells.Rendering.SheetRender sr = new Aspose.Cells.Rendering.SheetRender(workbook.Worksheets[0], imgOptions);
					int srPageCount = sr.PageCount;
					createZip = srPageCount > 1;
				}
			}



			string outfileName = Path.GetFileNameWithoutExtension(fileName) + outFileExtension;
            string outPath = "";

			string zipOutFolder = Aspose.App.Live.Demos.UI.Config.Configuration.OutputDirectory + guid;
            string zipOutfileName, zipOutPath;
            if (string.IsNullOrEmpty(zipFileName))
            {
                zipOutfileName = guid + ".zip";
				zipOutPath = Aspose.App.Live.Demos.UI.Config.Configuration.OutputDirectory + zipOutfileName;
            }
            else
            {
                var guid2 = Guid.NewGuid().ToString();
                outFolder = guid2;
                zipOutfileName = zipFileName + ".zip";
				zipOutPath = Aspose.App.Live.Demos.UI.Config.Configuration.OutputDirectory + guid2;
                Directory.CreateDirectory(zipOutPath);
                zipOutPath += "/" + zipOutfileName;
            }

            if (createZip)
            {
                outfileName = Path.GetFileNameWithoutExtension(fileName) + outFileExtension;
                outPath = zipOutFolder + "/" + outfileName;
                Directory.CreateDirectory(zipOutFolder);
            }
            else
            {
                outFolder = guid;
				outPath = Aspose.App.Live.Demos.UI.Config.Configuration.OutputDirectory +  outFolder;
                Directory.CreateDirectory(outPath);

                outPath += "/" + outfileName;
            }

            string statusValue = "OK";
            int statusCodeValue = 200;

            try
            {
                action(fileName, outPath, zipOutFolder);

                if (createZip)
                {
                    ZipFile.CreateFromDirectory(zipOutFolder, zipOutPath);
                    Directory.Delete(zipOutFolder, true);
                    outfileName = zipOutfileName;
                }

				if (deleteSourceFolder)
				{
					System.GC.Collect();
					System.GC.WaitForPendingFinalizers();
					Directory.Delete(sourceFolder, true);
				}

            }
            catch (Exception ex)
            {
                statusCodeValue = 500;
                statusValue = "500 " + ex.Message;
               
            }
            return new Response
            {
				FileName = outfileName,
				FolderName = outFolder,
				Status = statusValue,
				StatusCode = statusCodeValue,
			};
        }
		///<Summary>
		/// Process
		///</Summary>
		/// <param name="controllerName"></param>
		/// <param name="fileName"></param>
		/// <param name="folderName"></param>
		/// <param name="productName"></param>
		/// <param name="productFamily"></param>
		/// <param name="methodName"></param>
		/// <param name="action"></param>
		protected async Task<Response> Process(string controllerName, string fileName, string folderName, string productName, string productFamily, string methodName, inFileActionDelegate action)
        {           
            string tempFileName = fileName;
			string sourceFolder = Aspose.App.Live.Demos.UI.Config.Configuration.WorkingDirectory + folderName;
            fileName = sourceFolder + "/" + fileName;

            string statusValue = "OK";
            int statusCodeValue = 200;

            try
            {
                action(fileName);

                //Directory.Delete(sourceFolder, true);                

            }
            catch (Exception ex)
            {
                statusCodeValue = 500;
                statusValue = "500 " + ex.Message;
               
            }
            return await Task.FromResult(new Response
            {
                Status = statusValue,
                StatusCode = statusCodeValue,
            });
        }
		
    }
}
