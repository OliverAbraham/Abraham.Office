using System.Collections.Generic;

namespace Abraham.MicrosoftOfficeInterop.Fw461
{
	public class EmailParameters
	{
		public string		Subject     { get; set; }
		public string		Body        { get; set; }
		public string		Sender      { get; set; }
		public string		To          { get; set; }
		public string		Cc          { get; set; }
		public string		Bcc         { get; set; }
		public List<string> Attachments { get; set; }
	}
}
