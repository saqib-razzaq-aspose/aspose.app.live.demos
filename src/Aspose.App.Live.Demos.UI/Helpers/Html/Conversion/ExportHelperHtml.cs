using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Html.Rendering;

namespace Aspose.App.Live.Demos.UI.Helpers.Html.Conversion
{
	public class ExportHelperHtml : ExportHelper
	{
		public ExportHelperHtml(ExportFormat format) : base(format)
		{ }

		protected override void Render(string sourcePath, IDevice device)
		{
			using (Aspose.Html.HTMLDocument html_document = new Aspose.Html.HTMLDocument(sourcePath))
			{
				HtmlRenderer renderer = new HtmlRenderer();
				renderer.Render(device, html_document);
			}
		}
	}
}
