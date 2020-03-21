using Aspose.App.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using System.IO;
using System.Threading.Tasks;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	public partial class AsposeEmailConversionController
	{
		Response ConvertPst(string fileName, string folderName, string outputType)
		{
			switch (outputType)
			{
				case "eml": return ConvertPstToEml(fileName, folderName);
				case "msg": return ConvertPstToMsg(fileName, folderName);
				case "mbox": return ConvertPstToMbox(fileName, folderName);
				case "ost": return ConvertPstToOst(fileName, folderName);
				case "pst": return ReturnSame(fileName, folderName);
				case "mht": return ConvertPstToMht(fileName, folderName);
				case "html": return ConvertPstToHtml(fileName, folderName);
				case "svg": return ConvertPstToSvg(fileName, folderName);
				case "tiff": return ConvertPstToTiff(fileName, folderName);
				case "jpg": return ConvertPstToJpg(fileName, folderName);
				case "bmp": return ConvertPstToBmp(fileName, folderName);
				case "png": return ConvertPstToPng(fileName, folderName);
				default:
					return new Response
					{
						FileName = null,
						Status = $"Output type not supported {outputType.ToUpperInvariant()}",
						StatusCode = 500
					};
			}
		}

		Response ConvertPstToMsg(string fileName, string folderName)
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

		Response ConvertPstToMht(string fileName, string folderName)
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

		Response ConvertPstToHtml(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Html);
		}

		Response ConvertPstToSvg(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Svg);
		}

		Response ConvertPstToTiff(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Tiff);
		}

		Response ConvertPstToJpg(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Jpeg);
		}

		Response ConvertPstToBmp(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Bmp);
		}

		Response ConvertPstToPng(string fileName, string folderName)
		{
			return SaveOstOrPstAsDocument(fileName, folderName, Words.SaveFormat.Png);
		}

		Response ConvertPstToEml(string fileName, string folderName)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outputFolderPath)
			{
				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					int i = 0;
					HandleFolderAndSubfolders(mapiMessage =>
					{
						mapiMessage.Save(Path.Combine(outputFolderPath, "Message" + i++ + ".eml"), SaveOptions.DefaultEml);
					}, personalStorage.RootFolder, new MailConversionOptions());
				}
			});
		}

		///<Summary>
		/// ConvertPstToOst method to convert pst to ost file
		///</Summary>
		public  Response ConvertPstToOst(string fileName, string folderName)
		{
			return  ProcessTask(fileName, folderName, ".ost", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				using (var personalStorage = PersonalStorage.FromFile(inFilePath))
				{
					personalStorage.SaveAs(outPath, FileFormat.Ost);
				}
			});
		}

		///<Summary>
		/// ConvertPstToMbox method to convert pst to mbox file
		///</Summary>
		public  Response ConvertPstToMbox(string fileName, string folderName)
		{
			return  ProcessTask(fileName, folderName, ".mbox", false, false, delegate (string inFilePath, string outPath, string zipOutFolder)
			{
				var options = new MailConversionOptions();

				System.IO.File.Delete(outPath);
				using (FileStream writeStream = new FileStream(outPath, FileMode.Create))
				{
					using (MboxrdStorageWriter writer = new MboxrdStorageWriter(writeStream, false))
					{
						using (var pst = PersonalStorage.FromFile(inFilePath))
						{
							HandleFolderAndSubfolders(mapiMessage =>
							{
								var msg = mapiMessage.ToMailMessage(options);
								writer.WriteMessage(msg);
							}, pst.RootFolder, options);
						}
					}
				}
			});
		}
	}
}
