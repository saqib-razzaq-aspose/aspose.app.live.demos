using System.Web.Http;
using System.Threading.Tasks;
using Aspose.App.Live.Demos.UI.Models;
using System.Data;
using System.IO;


namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposeConversion class to call conversion method based on product name
	///</Summary>
	public class AsposeConversion : ApiBase
    {
		///<Summary>
		/// ConvertFile method to call conversion Controller based on product name
		///</Summary>
	
        public static Response ConvertFile(string fileName, string folderName, string outputType, string productName)
        {
			switch (productName)
            {
				case "words":
					AsposeWordsConversion wordsConversion = new AsposeWordsConversion();
					return  wordsConversion.ConvertFile(fileName, folderName, outputType);
				case "email":
					AsposeEmailConversionController emailController = new AsposeEmailConversionController();
					return  emailController.ConvertFile(fileName, folderName, outputType);
				case "cells":
                    AsposeCellsConversion cellsConversion = new AsposeCellsConversion();
                    return  cellsConversion.ConvertFile(fileName, folderName, outputType);
				case "slides":
					AsposeSlidesConversion asposeSlidesConversion = new AsposeSlidesConversion();
					return asposeSlidesConversion.ConvertFile(fileName, folderName, outputType);
				case "pdf":
					AsposePdfConversion asposePdfConversion = new AsposePdfConversion();
					return  asposePdfConversion.ConvertFile(fileName, folderName, outputType);
				case "imaging":
					AsposeImagingConversion imagingConversion = new AsposeImagingConversion();
					return  imagingConversion.ConvertFile(fileName, folderName, outputType);
				case "html":
					AsposeHtmlConversion htmlConversion = new AsposeHtmlConversion();
					return  htmlConversion.ConvertFile(fileName, folderName, outputType);
				case "tasks":
					AsposeTasksConversion tasksConversion = new AsposeTasksConversion();
					return  tasksConversion.ConvertFile(fileName, folderName, outputType);
				case "diagram":
					AsposeDiagramConversion diagramConversion = new AsposeDiagramConversion();
					return  diagramConversion.ConvertFile(fileName, folderName, outputType);
				case "note":
					AsposeNoteConversion noteConversion = new AsposeNoteConversion();
					return  noteConversion.ConvertFile(fileName, folderName, outputType);
				case "cad":
					AsposeCadConversion cadConversion = new AsposeCadConversion();
					return cadConversion.ConvertFile(fileName, folderName, outputType);
				case "gis":
					AsposeGisConversion gisConversion = new AsposeGisConversion();
					return gisConversion.ConvertFile(fileName, folderName, outputType);
				case "3d":
					Aspose3dConversion threeDConversion = new Aspose3dConversion();
					return threeDConversion.ConvertFile(fileName, folderName, outputType);

				case "psd":
					AsposePSDConversion psdConversion = new AsposePSDConversion();
					return psdConversion.ConvertFile(fileName, folderName, outputType);
				case "page":
					AsposePageConversion pageConversion = new AsposePageConversion();
					return  pageConversion.ConvertFile(fileName, folderName, outputType);
			}

            return new Response
            {
                FileName = null,
                Status = "Method not found",
                StatusCode = 500
            };

        }
		///<Summary>
		/// Convert Md File to PDF
		///</Summary>
	
        public async Task<Response> ConvertMdFile(string fileName, string folderName, string outputType, string productName)
        {
			// Commentend for Live Demos
            //switch (productName)
            //{                
            //    case "pdf":
            //        AsposePdfConversionController pdfController = new AsposePdfConversionController();
            //        return await pdfController.ConvertMarkownToPdf(fileName, folderName, outputType);
            //    case "html":
            //        AsposeHtmlConversionController htmlController = new AsposeHtmlConversionController();
            //        return await htmlController.ConvertMarkdownToHtml(fileName, folderName, outputType);
            //}

            return await Task.FromResult(new Response
            {
                FileName = null,
                Status = "Controller not found",
                StatusCode = 500
            });

        }
		///<Summary>
		/// Convert LaTeX File to PDF 
		///</Summary>
		[HttpGet]
		[ActionName("ConvertLatexToPdf")]
		public async Task<Response> ConvertLatexFile(string fileName, string folderName, string outputType, string productName)
		{
			// Commented for Live Demos
			//switch (productName)
			//{
			//	case "pdf":
			//		AsposePdfConversionController pdfController = new AsposePdfConversionController();
			//		return await pdfController.ConvertLatexToPdf(fileName, folderName, outputType);			
			//}

			return await Task.FromResult(new Response
			{
				FileName = null,
				Status = "Controller not found",
				StatusCode = 500
			});

		}
	}
}
