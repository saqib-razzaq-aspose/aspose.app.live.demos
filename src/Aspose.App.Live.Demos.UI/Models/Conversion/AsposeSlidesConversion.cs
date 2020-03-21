using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Diagnostics;
using Aspose.App.Live.Demos.UI.Services;
using Aspose.App.Live.Demos.UI.Helpers;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposeSlidesConversion class to convert slide files to different format
	///</Summary>
	public class AsposeSlidesConversion : ApiBase
	{
		///<Summary>
		/// ConvertFile method
		///</Summary>
		public Response ConvertFile(string fileName, string folderName, string outputType)
		{

			var pathProcessor = new PathProcessor(folderName, fileName, fileName != null);

			var result = Conversion(
						pathProcessor.DefaultSourceFile,
						pathProcessor.OutFolder, outputType

					);
			FileSafeResult fileSafeResult = new FileSafeResult();
			if (result == null)
			{
				fileSafeResult = pathProcessor.GetResultZipped();
			}
			else
			{
				fileSafeResult = pathProcessor.GetResult(Path.GetFileName(result));
			}

			return new Response
			{
				FileName = fileSafeResult.FileName,
				FolderName = fileSafeResult.id,
				Status = "OK",
				StatusCode = 200,
			};


		}
		public string Conversion(
			string sourceFile,
			string outFolder,
			string format
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				var fileName = Path.GetFileNameWithoutExtension(sourceFile);
				var outOneFile = Path.Combine(outFolder, $"{fileName}.{format}");

				switch (format)
				{
					case "odp":
					case "otp":
					case "pptx":
					case "pptm":
					case "potx":
					case "ppt":
					case "pps":
					case "ppsm":
					case "pot":
					case "potm":
					case "pdf":
					case "xps":
					case "ppsx":
					case "tiff":
					case "html":
					case "swf":
						var slidesFormat = format.ToString().ParseEnum<SaveFormat>();
						presentation.Save(outOneFile, slidesFormat);
						return outOneFile;

					case "txt":
						var lines = new List<string>();
						foreach (var slide in presentation.Slides)
						{
							foreach (var shp in slide.Shapes)
							{
								if (shp.Placeholder != null)
									lines.Add(((AutoShape)shp).TextFrame.Text);
							}

							var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

							if (!string.IsNullOrEmpty(notes))
								lines.Add(notes);
						}
						System.IO.File.WriteAllLines(outOneFile, lines);
						return outOneFile;

					case "doc":
					case "docx":
						using (var stream = new MemoryStream())
						{
							presentation.Save(stream, SaveFormat.Html);
							stream.Flush();
							stream.Seek(0, SeekOrigin.Begin);

							var doc = new Words.Document(stream);
							switch (format)
							{
								case "doc":
									doc.Save(outOneFile, Words.SaveFormat.Doc);
									break;

								case "docx":
									doc.Save(outOneFile, Words.SaveFormat.Docx);
									break;

								default:
									throw new ArgumentException($"Unknown format {format}");
							}
						}
						return outOneFile;

					case "bmp":
					case "jpeg":
					case "png":
					case "emf":
					case "wmf":
					case "gif":
					case "exif":
					case "ico":
						ImageFormat GetImageFormat(string f)
						{
							switch (format)
							{
								case "bmp":
									return ImageFormat.Bmp;
								case "jpeg":
									return ImageFormat.Jpeg;
								case "png":
									return ImageFormat.Png;
								case "emf":
									return ImageFormat.Wmf;
								case "wmf":
									return ImageFormat.Wmf;
								case "gif":
									return ImageFormat.Gif;
								case "exif":
									return ImageFormat.Emf;
								case "ico":
									return ImageFormat.Icon;
								default:
									throw new ArgumentException($"Unknown format {format}");
							}
						}

						///var size = presentation.SlideSize.Size;

						for (var i = 0; i < presentation.Slides.Count; i++)
						{
							var slide = presentation.Slides[i];
							var outFile = Path.Combine(outFolder, $"{i}.{format}");
							using (var bitmap = slide.GetThumbnail(1, 1))// (new Size((int)size.Width, (int)size.Height)))
								bitmap.Save(outFile, GetImageFormat(format));
						}

						return null;

					case "svg":
						var svgOptions = new SVGOptions
						{
							PicturesCompression = PicturesCompression.DocumentResolution
						};

						for (var i = 0; i < presentation.Slides.Count; i++)
						{
							var slide = presentation.Slides[i];
							var outFile = Path.Combine(outFolder, $"{i}.{format}");
							using (var stream = new FileStream(outFile, FileMode.CreateNew))
								slide.WriteAsSvg(stream, svgOptions);
						}

						return null;

					default:
						throw new ArgumentException($"Unknown format {format}");
				}
			}
		}
		
	}
}
