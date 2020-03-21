using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Html.Rendering;
using Aspose.Html.Dom.Svg;

namespace Aspose.App.Live.Demos.UI.Helpers.Html.Conversion
{
	public class ExportHelperSvg : ExportHelper
	{
		public ExportHelperSvg(ExportFormat format) : base(format)
		{ }

		protected override void Render(string sourcePath, IDevice device)
		{
			using (SVGDocument document = new SVGDocument(sourcePath))
			{
				var renderer = new SvgRenderer();
				renderer.Render(device, document);
			}
		}
	}
}
