using System.Diagnostics;

namespace Abraham.Office
{
	public class DocxToPdfConverter
	{
		public static void Convert(string inputFilename, string pdfOutputFilename, bool hide = true, bool exit = true)
		{
			var tempFilename = Path.GetFileNameWithoutExtension(inputFilename) + ".pdf";
			if (File.Exists(tempFilename))
				File.Delete(tempFilename);

			var program = @"C:\Program Files (x86)\NCH Software\Doxillion\doxillion.exe";
			var arguments = $"-convert -format PDF -overwrite ALWAYS ";
		                  
			if (hide)
				arguments += "-hide ";
			if (exit)
				arguments += "-exit ";

			arguments += $"\"{Path.GetFullPath(inputFilename)}\"";

			Process.Start(program, arguments);

			var circuitBreaker = 60;
			while (!File.Exists(tempFilename))
			{
				Thread.Sleep(500);
				if (--circuitBreaker <= 0)
					break;
			}

			if (File.Exists(tempFilename) && tempFilename != pdfOutputFilename)
				File.Copy(tempFilename, pdfOutputFilename, true);
		}
	}
}
