using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Html.Rendering;

namespace Aspose.App.Live.Demos.UI.Helpers.Html.Conversion
{
	public class ExportHelperMhtml : ExportHelper
	{
		public ExportHelperMhtml(ExportFormat format) : base(format)
		{ }

		protected override void Render(string sourcePath, IDevice device)
		{
			using(FileStream fstr = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
			{
				var renderer = new MhtmlRenderer();
				renderer.Render(device, fstr);
			}
		}
	}
}
