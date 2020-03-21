using Aspose.App.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Aspose.App.Live.Demos.UI.Controllers
{
	public class HomeController : BaseController
	{
	
		public override string Product => (string)RouteData.Values["productname"];
		

		public ActionResult Index()
		{
			ViewBag.PageTitle = "Free File Format Apps - Process MS Word | PDF | Excel | PPT online";
			ViewBag.MetaDescription = "Free Apps to Read, Manipulate, Convert MS Word, PDF, Excel, PowerPoint, Visio, MS Project, OneNote, Email, MSG, Barcode, CAD, 3D, GIS, HTML file formats online";
			var model = new LandingPageModel(this)

			{
				Product = Product
			};

			return View(model);
		}		

		public ActionResult Default()
		{
			ViewBag.PageTitle = "Free File Format Apps - Process MS Word | PDF | Excel | PPT online";
			ViewBag.MetaDescription = "Free Apps to Read, Manipulate, Convert MS Word, PDF, Excel, PowerPoint, Visio, MS Project, OneNote, Email, MSG, Barcode, CAD, 3D, GIS, HTML file formats online";
			var model = new LandingPageModel(this);

			return View(model);
		}
	}
}
