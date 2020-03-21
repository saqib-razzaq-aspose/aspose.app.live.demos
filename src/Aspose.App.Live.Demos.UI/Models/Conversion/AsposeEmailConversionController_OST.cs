using Aspose.App.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using System.IO;
using System.Threading.Tasks;
using System.Web.Http;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	public partial class AsposeEmailConversionController
	{
		Response ConvertOst(string fileName, string folderName, string outputType)
		{
			switch (outputType)
			{
				case "eml": return ConvertOstToEml(fileName, folderName);
				case "msg": return ConvertOstToMsg(fileName, folderName);
				case "mbox": return ConvertOstToMbox(fileName, folderName);
				case "ost": return ReturnSame(fileName, folderName);
				case "pst": return ConvertOSTToPST(fileName, folderName, "");
				case "mht": return ConvertOstToMht(fileName, folderName);
				case "html": return ConvertOstToHtml(fileName, folderName);
				case "svg": return ConvertOstToSvg(fileName, folderName);
				case "tiff": return ConvertOstToTiff(fileName, folderName);
				case "jpg": return ConvertOstToJpg(fileName, folderName);
				case "bmp": return ConvertOstToBmp(fileName, folderName);
				case "png": return ConvertOstToPng(fileName, folderName);
				default:
					return new Response
					{
						FileName = null,
						Status = $"Output type not supported {outputType.ToUpperInvariant()}",
						StatusCode = 500
					};
			}
		}

		Response ConvertOstToPng(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Png);
		}

		Response ConvertOstToBmp(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Bmp);
		}

		Response ConvertOstToJpg(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Jpeg);
		}

		Response ConvertOstToTiff(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Tiff);
		}

		Response ConvertOstToSvg(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Svg);
		}

		Response ConvertOstToHtml(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Html);
		}

		Response ConvertOstToMbox(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					var options = new MailConversionOptions();

					using (FileStream writeStream = new FileStream(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + ".mbox"), FileMode.Create))
					{
						using (MboxrdStorageWriter writer = new MboxrdStorageWriter(writeStream, false))
						{
							HandleFolderAndSubfolders(mapiMessage =>
							{
								var msg = mapiMessage.ToMailMessage(options);
								writer.WriteMessage(msg);
							}, personalStorage.RootFolder, options);
						}
					}
				}
			});
		}

		Response ConvertOstToMht(string fileName, string folderName)
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

				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					int i = 0;
					HandleFolderAndSubfolders(mapiMessage =>
					{
						mapiMessage.Save(Path.Combine(outputFolderPath, "Message" + i++ + ".mht"), mhtSaveOptions);
					}, personalStorage.RootFolder, new MailConversionOptions());
				}
			});
		}

		Response ConvertOstToMsg(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					int i = 0;
					HandleFolderAndSubfolders(mapiMessage =>
					{
						mapiMessage.Save(Path.Combine(outputFolderPath, "Message" + i++ + ".msg"), SaveOptions.DefaultMsgUnicode);
					}, personalStorage.RootFolder, new MailConversionOptions());
				}
			});
		}

		Response ConvertOstToEml(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					int i = 0;
					HandleFolderAndSubfolders(mapiMessage =>
					{
						mapiMessage.Save(Path.Combine(outputFolderPath, "Message" + i++ + ".eml"), SaveOptions.DefaultEml );
					}, personalStorage.RootFolder, new MailConversionOptions());
				}
			});
		}

		///<Summary>
		/// ConvertOSTToPST method to convert ost to pst file
		///</Summary>

		public  Response ConvertOSTToPST(string fileName, string folderName, string userEmail)
		{
			return  ProcessTask(fileName, folderName, ".pst", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				using (var personalStorage = PersonalStorage.FromFile(inFilePath))
				{
					personalStorage.SaveAs(outPath, FileFormat.Pst);
				}
			});
		}
	}
}
