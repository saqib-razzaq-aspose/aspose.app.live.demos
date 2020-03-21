using Aspose.App.Live.Demos.UI.Config;
using Aspose.App.Live.Demos.UI.Models;
using Aspose.App.Live.Demos.UI.Services;
using Aspose.App.Live.Demos.UI.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;

namespace Aspose.App.Live.Demos.UI.Controllers
{
	public abstract class BaseController : Controller
	{
		///// <summary>
		///// Upload one or more Documents and save to temporary local folder
		///// </summary>
		///// <param name="inputType"></param>
		///// <returns></returns>
		//protected async string[] UploadDocuments(string inputType)
		//{
		//	var _inputType = inputType.ToLower();
		//	try
		//	{
		//		var tmpFolderName = Guid.NewGuid().ToString();
		//		var pathProcessor = new PathProcessor(tmpFolderName);
		//		var uploadProvider = new MultipartFormDataStreamProviderSafe(pathProcessor.SourceFolder);
		//		  Request.Content.ReadAsMultipartAsync(uploadProvider);
		//		return uploadProvider.FileData.Select(x => x.LocalFileName).ToArray();
		//	}
		//	catch (Exception ex)
		//	{

		//		return new string[0];
		//	}
		//}
		/// <summary>
		/// Response when uploaded files exceed the limits
		/// </summary>
		protected Response BadDocumentResponse = new Response()
		{
			Status = "Some of your documents are corrupted",
			StatusCode = 500
		};
		

		public abstract string Product { get; }

		protected override void OnActionExecuted(ActionExecutedContext ctx)
		{
			base.OnActionExecuted(ctx);			
			ViewBag.Title = ViewBag.Title ?? Resources["ApplicationTitle"];
			ViewBag.MetaDescription = ViewBag.MetaDescription ?? "Save time and software maintenance costs by running single instance of software, but serving multiple tenants/websites. Customization available for Joomla, Wordpress, Discourse, Confluence and other popular applications.";
		}

		private AsposeAppContext _atcContext;
		

		/// <summary>
		/// Main context object to access all the dcContent specific context info
		/// </summary>
		public AsposeAppContext AsposeToolsContext
		{
			get
			{
				if (_atcContext == null) _atcContext = new AsposeAppContext(HttpContext.ApplicationInstance.Context);
				return _atcContext;
			}
		}

		private Dictionary<string, string> _resources;

		/// <summary>
		/// key/value pair containing all the error messages defined in resources.xml file
		/// </summary>
		public Dictionary<string, string> Resources
		{
			get
			{
				if (_resources == null) _resources = AsposeToolsContext.Resources;
				return _resources;
			}
		}

		protected bool IsValidEmail(string email)
		{
			try
			{
				var addr = new System.Net.Mail.MailAddress(email);
				return addr.Address == email;
			}
			catch
			{
				return false;
			}
		}

		
	}
}
