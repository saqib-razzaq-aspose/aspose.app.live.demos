using Aspose.App.Live.Demos.UI.Helpers;
using Aspose.App.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Words;
using System.IO;
using System.Threading.Tasks;
using System.Web.Http;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	public partial class AsposeEmailConversionController
	{
		Response ConvertEmlOrMsg(string fileName, string folderName, string outputType)
		{
			switch (outputType)
			{
				case "eml": return ConvertMIMEMessageToEML(fileName, folderName, "");
				case "msg": return ConvertEMLToMSG(fileName, folderName, "");
				case "mbox": return ConvertMailToMbox(fileName, folderName);
				case "ost": return ConvertMailToOst(fileName, folderName);
				case "pst": return ConvertMailToPst(fileName, folderName);
				case "mht": return ConvertEMLToMHT(fileName, folderName, "");
				case "html": return ConvertEmailToHtml(fileName, folderName);
				case "svg":
				case "tiff": return ConvertEmailToSingleImage(fileName, folderName, outputType);
				case "jpg":
				case "bmp":
				case "png": return ConvertEmailToImages(fileName, folderName, outputType);
				default:
					return new Response
					{
						FileName = null,
						Status = $"Output type not supported {outputType.ToUpperInvariant()}",
						StatusCode = 500
					};
			}
		}

		
		public Response ConvertEMLToMSG(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".msg", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load mail message
				var message = MailMessage.Load(inFilePath, new EmlLoadOptions());
				// Save as MSG
				message.Save(outPath, SaveOptions.DefaultMsgUnicode);
			});
		}
		///<Summary>
		/// ConvertMIMEMessageToEML method to convert mime message to eml 
		///</Summary>
	
		public Response ConvertMIMEMessageToEML(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".eml", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load mail message
				var message = MapiHelper.GetMailMessageFromFile(inFilePath);
				// Save as EML
				message.Save(outPath, SaveOptions.DefaultEml);
			});
		}

		///<Summary>
		/// ConvertEMLToMHT method to convert eml to mht
		///</Summary>

		public Response ConvertEMLToMHT(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".mht", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				// Load mail message
				var mail = MapiHelper.GetMailMessageFromFile(inFilePath);

				// Save as mht with header
				var mhtSaveOptions = new MhtSaveOptions
				{
					//Specify formatting options required
					//Here we are specifying to write header informations to output without writing extra print header
					//and the output headers should display as the original headers in message
					MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.HideExtraPrintHeader | MhtFormatOptions.DisplayAsOutlook,
					// Check the body encoding for validity. 
					CheckBodyContentEncoding = true
				};

				mail.Save(outPath, mhtSaveOptions);
			});
		}
		///<Summary>
		/// ConvertEmailToHtml method to email to html
		///</Summary>
		public Response ConvertEmailToHtml(string fileName, string folderName)
		{
			return  ProcessTask(fileName, folderName, ".html", true, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				var eml = MailMessage.Load(inFilePath, new EmlLoadOptions());

				var htmlSaveOptions = new HtmlSaveOptions
				{
					HtmlFormatOptions = HtmlFormatOptions.WriteHeader | HtmlFormatOptions.DisplayAsOutlook,
					CheckBodyContentEncoding = true
				};

				eml.Save(outPath, htmlSaveOptions);
			});
		}
		///<Summary>
		/// ConvertEmailToImages method to convert email to messages
		///</Summary>
		public Response ConvertEmailToImages(string fileName, string folderName, string outputType)
		{
			if (outputType.Equals("bmp") || outputType.Equals("jpg") || outputType.Equals("png"))
			{
				var format = Aspose.Words.SaveFormat.Bmp;

				if (outputType.Equals("jpg"))
				{
					format = Aspose.Words.SaveFormat.Jpeg;
				}
				else if (outputType.Equals("png"))
				{
					format = Aspose.Words.SaveFormat.Png;
				}

				return  ProcessTask(fileName, folderName, "." + outputType, true, true, delegate (string inFilePath, string outPath, string zipOutFolder)
				{
					var msg = MapiHelper.GetMailMessageFromFile(inFilePath);

					var msgStream = new MemoryStream();
					msg.Save(msgStream, SaveOptions.DefaultMhtml);
					msgStream.Position = 0;

					var options = new Aspose.Words.Saving.ImageSaveOptions(format);
					options.PageCount = 1;

					var outfileName = Path.GetFileNameWithoutExtension(fileName) + "_{0}";
					var doc = new Document(msgStream);
					var pageCount = doc.PageCount;

					for (int i = 0; i < doc.PageCount; i++)
					{
						if (pageCount > 1)
						{
							outPath = zipOutFolder + "/" + outfileName;
							options.PageIndex = i;
							doc.Save(string.Format(outPath, (i + 1) + "." + outputType), options);
						}
						else
						{
							outPath = zipOutFolder + "/" + Path.GetFileNameWithoutExtension(fileName);
							options.PageIndex = i;
							doc.Save(outPath + "." + outputType, options);
						}

					}

				});
			}

			return new Response
			{
				FileName = null,
				Status = "Output type not found",
				StatusCode = 500
			};
		}

		///<Summary>
		/// ConvertEmailToSingleImage method to convert email to single image
		///</Summary>
		public Response ConvertEmailToSingleImage(string fileName, string folderName, string outputType)
		{
			if (outputType.Equals("tiff") || outputType.Equals("svg"))
			{
				var format = Aspose.Words.SaveFormat.Tiff;

				if (outputType.Equals("svg"))
				{
					format = Aspose.Words.SaveFormat.Svg;
				}

				return ProcessTask(fileName, folderName, "." + outputType, false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
				{
					var msg = MapiHelper.GetMailMessageFromFile(inFilePath);

					var msgStream = new MemoryStream();
					msg.Save(msgStream, SaveOptions.DefaultMhtml);
					msgStream.Position = 0;

					var msgDocument = new Document(msgStream);
					msgDocument.Save(outPath, format);
				});
			}

			return new Response
			{
				FileName = null,
				Status = "Output type not found",
				StatusCode = 500
			};
		}

		Response ConvertMailToPst(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inFilePath, string outPath)
			{
				var msg = MapiHelper.GetMailMessageFromFile(inFilePath);
				var pstFileName = Path.Combine(outPath, Path.GetFileNameWithoutExtension(fileName) + ".pst");

				using (var personalStorage = PersonalStorage.Create(pstFileName, FileFormatVersion.Unicode))
				{
					var inbox = personalStorage.RootFolder.AddSubFolder("Inbox");

					inbox.AddMessage(MapiMessage.FromMailMessage(msg, MapiConversionOptions.UnicodeFormat));
				}
			});
		}

		Response ConvertMailToOst(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inFilePath, string outPath)
			{
				var msg = MapiHelper.GetMailMessageFromFile(inFilePath);
				var temporaryPstFileName = Path.Combine(outPath, Path.GetFileNameWithoutExtension(fileName) + ".pst");

				try
				{
					using (var personalStorage = PersonalStorage.Create(temporaryPstFileName, FileFormatVersion.Unicode))
					{
						var inbox = personalStorage.RootFolder.AddSubFolder("Inbox");

						inbox.AddMessage(MapiMessage.FromMailMessage(msg, MapiConversionOptions.UnicodeFormat));

						var ostFileName = Path.Combine(outPath, Path.GetFileNameWithoutExtension(fileName) + ".ost");
						personalStorage.SaveAs(ostFileName, FileFormat.Ost);
					}
				}
				finally
				{
					if (System.IO.File.Exists(temporaryPstFileName))
						System.IO.File.Delete(temporaryPstFileName);
				}
			});
		}

		Response ConvertMailToMbox(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inFilePath, string outPath)
			{
				var msg = MapiHelper.GetMailMessageFromFile(inFilePath);

				using (var writeStream = new FileStream(Path.Combine(outPath, Path.GetFileNameWithoutExtension(fileName) + ".mbox"), FileMode.Create))
				{
					using (var writer = new MboxrdStorageWriter(writeStream, false))
					{
						writer.WriteMessage(msg);
					}
				}
			});
		}
	}
}
