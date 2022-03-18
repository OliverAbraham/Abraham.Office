using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;

namespace Abraham.MicrosoftOfficeInterop.Fw461
{
	/// <summary>
	/// Helper application to open microsoft outlooks new Email window, from the Outlook 2019 thats installed on the local machine
	/// </summary>
	internal class Program
	{
		[DllImport("user32.dll")]
		static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		static void Main(string[] args)
		{
			HideTheWindow();
			OpenUpMicrosoftOutlookNewEmail(args);
		}

		private static void HideTheWindow()
		{
			try
			{
				IntPtr h = Process.GetCurrentProcess().MainWindowHandle;
				ShowWindow(h, 0);
			}
			catch (System.Exception ex)
			{
				Console.WriteLine(ex);
			}
		}

		private static void OpenUpMicrosoftOutlookNewEmail(string[] args)
		{
			try
			{
				OpenNewEmailWindow(args[0]);
			}
			catch (System.Exception ex)
			{
				Console.WriteLine(ex);
				Thread.Sleep(10 * 1000);
			}
		}

		private static void OpenNewEmailWindow(string argumentsBase64encoded)
		{
			var argumentBytes = Convert.FromBase64String(argumentsBase64encoded);
			var arguments = Encoding.UTF8.GetString(argumentBytes);

			Console.WriteLine($"Arguments: {arguments}");

			var parameters = JsonConvert.DeserializeObject<EmailParameters>(arguments);

			var subject     = parameters.Subject;
			var body        = parameters.Body;
			var sender      = parameters.Sender;
			var to          = parameters.To;
			var cc          = parameters.Cc;
			var bcc         = parameters.Bcc;
			var attachments = parameters.Attachments;

			var app = new Application();
			var msg = (MailItem)app.CreateItem(OlItemType.olMailItem);

			foreach (Account account in app.Explorers.Session.Accounts)
			{
				if (account.DisplayName == sender)
					msg.SendUsingAccount = account;
			}

			if (!string.IsNullOrEmpty(to))
				msg.To  = to;
			if (!string.IsNullOrEmpty(cc))
				msg.CC  = cc;
			if (!string.IsNullOrEmpty(bcc))
				msg.BCC = bcc;
			if (attachments != null)
			{
				foreach(var attachment in attachments)
					msg.Attachments.Add(attachment.Replace('\\', '/'), OlAttachmentType.olByValue, Type.Missing, Type.Missing);
			}

			msg.Subject = subject;
			msg.BodyFormat = OlBodyFormat.olFormatHTML;
			msg.HTMLBody = body;
			msg.Display(false); //In order to display it in modal inspector change the argument to true
			
			Console.WriteLine("fertig");
			Thread.Sleep(10*1000);
		}
	}
}
