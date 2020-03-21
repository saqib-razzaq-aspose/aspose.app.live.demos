using Aspose.App.Live.Demos.UI.Controllers;
using  Aspose.App.Live.Demos.UI.Models;
using System;
using System.Diagnostics;
using System.IO;


namespace Aspose.App.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// LoggableFileProcessor class to log information in database
	///</Summary>
	public abstract class LoggableFileProcessor : FileProcessor
    {
		
		///<Summary>
		/// Process
		///</Summary>
		/// <param name="inputFolderName"></param>
		/// <param name="inputFileName"></param>
		/// <param name="outFileName"></param>
		public override Response Process(string inputFolderName, string inputFileName, string outFileName = null)
        {
            var stackTrace = new StackTrace();

            string methodName = null;
            string controllerName = null;

            for (int i = 1; i < stackTrace.FrameCount; i++)
            {
                var method = stackTrace.GetFrame(i).GetMethod();

                if (method.DeclaringType != null && typeof(ApiBase).IsAssignableFrom(method.DeclaringType))
                {
                    methodName = method.Name;
                    controllerName = method.DeclaringType.Name;
                    break;
                }
                else
                {
                    methodName = method.Name;
                    controllerName = "NULL";
                }
            }

            var folderName = Path.GetDirectoryName(inputFileName);
            

            try
            {
                var resp = base.Process(inputFolderName, inputFileName, outFileName);
               
                return resp;
            }
            catch (Exception ex)
            {
                var resp = new Response()
                {
                    StatusCode = 500,
                    Status = "500 " + ex.Message
                };

                // Log error message to NLogging database
               
                return resp;
            }
        }
    }
}
