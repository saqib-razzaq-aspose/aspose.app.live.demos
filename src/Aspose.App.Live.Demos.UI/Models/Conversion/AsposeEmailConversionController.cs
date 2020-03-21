using Aspose.App.Live.Demos.UI.FileProcessing;
using Aspose.App.Live.Demos.UI.Models;
using Aspose.App.Live.Demos.UI.Models.Conversion;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using Aspose.Words;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using static Aspose.App.Live.Demos.UI.FileProcessing.CustomSingleOrZipFileProcessor;

namespace Aspose.App.Live.Demos.UI.Models.Conversion
{
	///<Summary>
	/// AsposeEmailConversionController class to convert email file to different format
	///</Summary>
	public partial class AsposeEmailConversionController : ApiBase
	{
		///<Summary>
		/// ConvertFile
		///</Summary>
		public Response ConvertFile(string fileName, string folderName, string outputType)
		{
			fileName = fileName.ToLowerInvariant();
			outputType = outputType.ToLowerInvariant();

			var ext = Path.GetExtension(fileName);

			/* All conversions table:
			 * 
			 * INPUTS   OUTPUTS
			 * 
			 * eml		eml
			 * msg		msg
			 * mbox		mbox
			 * ost		ost
			 * pst		pst
			 *			mht
			 *			html
			 *			svg
			 *			tiff
			 *			jpg
			 *			bmp
			 *			png
			 */

			switch (ext)
			{
				case ".eml": 
				case ".msg": return ConvertEmlOrMsg(fileName, folderName, outputType);
				case ".mbox": return ConvertMbox(fileName, folderName, outputType);
				case ".ost": return ConvertOst(fileName, folderName, outputType);
				case ".pst": return ConvertPst(fileName, folderName, outputType);
				default:
					return new Response
					{
						FileName = null,
						Status = $"Input type not supported {ext.ToUpperInvariant()}",
						StatusCode = 500
					};
			}
		}

		Response ReturnSame(string fileName, string folderName)
		{
			return new Response
			{
				FileName = fileName,
				FolderName = folderName,
				StatusCode = 200
			};
		}

		Response ProcessTask(string fileName, string folderName, string outFileExtension, bool createZip, bool checkNumberofPages, ActionDelegate action)
		{
			Aspose.App.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
			Aspose.App.Live.Demos.UI.Models.License.SetAsposeWordsLicense(); // for email to image conversion
			return Process(this.GetType().Name, fileName, folderName, outFileExtension, createZip, checkNumberofPages,  (new StackTrace()).GetFrame(5).GetMethod().Name, action);
		}

		Response ProcessTask(string fileName, string folderName, ProcessFileDelegate handler)
		{
			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomProcessMethod = handler
			};

			Aspose.App.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
			Aspose.App.Live.Demos.UI.Models.License.SetAsposeWordsLicense(); // for email to image conversion

			return processor.Process(folderName, fileName);
		}

		void HandleFolderAndSubfolders(Action<MapiMessage> handler, FolderInfo folderInfo, MailConversionOptions options)
		{
			foreach (MapiMessage mapiMessage in folderInfo.EnumerateMapiMessages())
			{
				handler(mapiMessage);
			}

			if (folderInfo.HasSubFolders == true)
			{
				foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
				{
					HandleFolderAndSubfolders(handler, subfolderInfo, options);
				}
			}
		}

		Response SaveOstOrPstAsDocument(string fileName, string folderName, SaveFormat format)
		{
			return ProcessTask(fileName, folderName, delegate (string inputFilePath, string outPath)
			{
				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					var msgStream = new MemoryStream();

					HandleFolderAndSubfolders(mapiMessage =>
					{
						mapiMessage.Save(msgStream, SaveOptions.DefaultMhtml);
					}, personalStorage.RootFolder, new MailConversionOptions());

					msgStream.Position = 0;
					SaveDocumentStreamToFolder(msgStream, Path.GetFileNameWithoutExtension(fileName), outPath, format);
				}
			});
		}

		void SaveDocumentStreamToFolder(Stream stream, string fileName, string outPath, SaveFormat format)
		{
			var document = new Document(stream);
			var shortFileName = Path.GetFileNameWithoutExtension(fileName);
			var formatExt = "." + format.ToString().ToLowerInvariant();

			switch (format)
			{
				//case SaveFormat.Svg:
				//case SaveFormat.Tiff:
				case SaveFormat.Png:
				case SaveFormat.Bmp:
				case SaveFormat.Emf:
				case SaveFormat.Jpeg:
				case SaveFormat.Gif:
					{
						var pageCount = document.PageCount;

						if (pageCount == 1)
						{
							document.Save(Path.Combine(outPath, shortFileName + formatExt), format);
							break;
						}

						var options = new Aspose.Words.Saving.ImageSaveOptions(format);
						options.PageCount = 1;

						for (int i = 0; i < document.PageCount; i++)
						{
							options.PageIndex = i;
							document.Save(Path.Combine(outPath, shortFileName + i + formatExt), format);
						}
						break;
					}
				default:
					document.Save(Path.Combine(outPath, shortFileName + formatExt), format);
					break;
			}
		}
	}
}
