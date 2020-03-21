using Aspose.App.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Storage;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Words;
using System.IO;
using System.Threading.Tasks;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	public partial class AsposeEmailConversionController
	{
		Response ConvertMbox(string fileName, string folderName, string outputType)
		{
			switch (outputType)
			{
				case "eml": return ConvertMboxToEml(fileName, folderName);
				case "msg": return ConvertMboxToMsg(fileName, folderName);
				case "mbox": return ReturnSame(fileName, folderName);
				case "ost": return ConvertMboxToOst(fileName, folderName);
				case "pst": return ConvertMboxToPst(fileName, folderName);
				case "mht": return ConvertMboxToMht(fileName, folderName);
				case "html": return ConvertMboxToHtml(fileName, folderName);
				case "svg": return ConvertMboxToSvg(fileName, folderName);
				case "tiff": return ConvertMboxToTiff(fileName, folderName);
				case "jpg": return ConvertMboxToJpg(fileName, folderName);
				case "bmp": return ConvertMboxToBmp(fileName, folderName);
				case "png": return ConvertMboxToPng(fileName, folderName);
				default:
					return new Response
					{
						FileName = null,
						Status = $"Output type not supported {outputType.ToUpperInvariant()}",
						StatusCode = 500
					};
			}
		}

		Response ConvertMboxToTiff(string fileName, string folderName)
		{
			return SaveMboxAsDocument(fileName, folderName, SaveFormat.Tiff, ".tiff");
		}

		Response ConvertMboxToJpg(string fileName, string folderName)
		{
			return SaveMboxAsDocument(fileName, folderName, SaveFormat.Jpeg, ".jpg");
		}

		Response ConvertMboxToBmp(string fileName, string folderName)
		{
			return SaveMboxAsDocument(fileName, folderName, SaveFormat.Bmp, ".bmp");
		}

		Response ConvertMboxToPng(string fileName, string folderName)
		{
			return SaveMboxAsDocument(fileName, folderName, SaveFormat.Png, ".png");
		}

		Response ConvertMboxToSvg(string fileName, string folderName)
		{
			return SaveMboxAsDocument(fileName, folderName, SaveFormat.Svg, ".svg");
		}

		Response ConvertMboxToHtml(string fileName, string folderName)
		{
			return SaveMboxAsDocument(fileName, folderName, SaveFormat.Html, ".html");
		}

		Response ConvertMboxToMht(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
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

				using (var reader = new MboxrdStorageReader(inputFilePath, false))
				{
					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
					{
						var message = reader.ReadNextMessage();
						message.Save(Path.Combine(outputFolderPath, "Message" + i + ".mht"), mhtSaveOptions);
					}
				}
			});
		}

		Response ConvertMboxToOst(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
				var shortFileName = Path.GetFileNameWithoutExtension(inputFilePath);
				var tempPath = Path.Combine(outputFolderPath, shortFileName+".pst");

				try
				{
					using (var pst = MailStorageConverter.MboxToPst(inputFilePath, tempPath))
						pst.SaveAs(Path.Combine(outputFolderPath, shortFileName + ".ost"), FileFormat.Ost);
				}
				finally
				{
					System.IO.File.Delete(tempPath);
				}
			});
		}

		///<Summary>
		/// ConvertMboxToPst method to convert mbox to pst file
		///</Summary>
		public Response ConvertMboxToPst(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, ".pst", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				MailStorageConverter.MboxToPst(inFilePath, outPath).Dispose();
			});
		}

		///<Summary>
		/// ConvertMboxToEml method to convert mbox to eml file
		///</Summary>
		Response ConvertMboxToEml(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
				using (var reader = new MboxrdStorageReader(inputFilePath, false))
				{
					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
					{
						var message = reader.ReadNextMessage();
						message.Save(Path.Combine(outputFolderPath, "Message" + i + ".eml"), SaveOptions.DefaultEml);
					}
				}
			});
		}

		///<Summary>
		/// ConvertMboxToMsg method to convert mbox to msg file
		///</Summary>
		Response ConvertMboxToMsg(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
				{
					using (var reader = new MboxrdStorageReader(inputFilePath, false))
					{
						for (int i = 0; i < reader.GetTotalItemsCount(); i++)
						{
							var message = reader.ReadNextMessage();
							message.Save(Path.Combine(outputFolderPath, "Message" + i + ".msg"), SaveOptions.DefaultMsgUnicode);
						}
					}
				});
		}

		Response SaveMboxAsDocument(string fileName, string folderName, SaveFormat format, string saveExtension)
		{
			return ProcessTask(fileName, folderName, saveExtension, false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				using (var reader = new MboxrdStorageReader(inFilePath, false))
				{
					var msgStream = new MemoryStream();

					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
					{
						var message = reader.ReadNextMessage();
						message.Save(msgStream, SaveOptions.DefaultMhtml);
					}

					msgStream.Position = 0;

					SaveDocumentStreamToFolder(msgStream, inFilePath, Directory.GetParent(outPath).FullName, format);
				}
			});
		}
	}
}
