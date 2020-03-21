using Aspose.App.Live.Demos.UI.Models.Conversion;
using Aspose.App.Live.Demos.UI.Models;
using Aspose.App.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.App.Live.Demos.UI.Controllers
{
	public class AsposeConversionController : BaseController
	{
		public override string Product => (string)RouteData.Values["product"];
		
		[HttpPost]
		public Response Conversion(string outputType, string productName)
		{
			Response response = null;
			var files = Request.Files;
			foreach (string fileName in Request.Files)
			{
				HttpPostedFileBase postedFile = Request.Files[fileName];

				if (postedFile != null)
				{
					var isFileUploaded = FileManager.UploadFile(postedFile);

					if ((isFileUploaded != null) && (isFileUploaded.FileName.Trim() != ""))
					{
						 response = AsposeConversion.ConvertFile(isFileUploaded.FileName, isFileUploaded.FolderId, outputType.ToLower().Replace(" ", ""), productName);

						if (response == null)
						{
							throw new Exception(Resources["APIResponseTime"]);
						}				

					}
				}

			}

			return response;			
				
		}

		

		public ActionResult Conversion()
		{
			var model = new ViewModel(this, "Conversion")
			{
				SaveAsComponent = true,
				SaveAsOriginal = false
			};

			return View(model);
		}
		

	}
}
