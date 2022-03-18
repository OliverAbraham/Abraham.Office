using Newtonsoft.Json;
using System.Diagnostics;
using System.Text;

namespace Abraham.MicrosoftOfficeInterop
{
	/// <summary>
	/// A class to open MS Outlook "new email" window with subject, body and attachment
	/// </summary>
	public class OfficeInterop
	{
		public static void OpenNewEmailWindow(EmailParameters parameters)
		{
			var serializedParameters = JsonConvert.SerializeObject(parameters);
			var base64Converted = Convert.ToBase64String(Encoding.UTF8.GetBytes(serializedParameters));
			
			var info = new ProcessStartInfo();
			info.FileName = "Abraham.MicrosoftOfficeInterop.Fw461.exe";
			info.Arguments = base64Converted;
			info.WindowStyle = ProcessWindowStyle.Minimized;
			Process.Start(info);
		}
	}
}