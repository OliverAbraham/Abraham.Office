using Abraham.MicrosoftOfficeInterop;
using Abraham.ShellFunctions;
using Abraham.ShellLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace Abraham.Office.Demo
{
	/// <summary>
	/// Demo application for Abraham.Office nuget package
	/// </summary>
	public partial class MainWindow : Window
	{
        private const string docxTemplateFilename = "Template.docx";
        private const string docxOutputFilename   = "GeneratedFile.docx";
		private const string pdfOutputFilename    = "GeneratedFile.pdf";
            

		public MainWindow()
		{
			InitializeComponent();
			textboxDate.Text = DateTime.Now.ToShortDateString();
		}

		private void Button_EditTemplate_Click(object sender, RoutedEventArgs e)
		{
			ExternalPrograms.OpenFileInEditor(Path.GetFullPath("Template.docx"));
		}

		private void Button_GenerateDocx_Click(object sender, RoutedEventArgs e)
		{
            var fields = new List<FieldValue>()
            { 
                new FieldValue("[INVOICEDATE]"  , textboxDate.Text),
                new FieldValue("[RECEIVER]"     , "Example John Doe"),
                new FieldValue("[STREET]"       , "Example street"),
                new FieldValue("[NUMBER]"       , "1234"),
                new FieldValue("[POSTAL]"       , "12345"),
                new FieldValue("[CITY]"         , "Example City"),
                new FieldValue("[INVOICENUMBER]", "123456"),
            };
            var tableCellNames = new List<string> { "[DATE]", "[ITEMTEXT]", "[PRICE1]", "[PRICE2]" };
            var tableRows = new List<TableRow>()
            { 
                new TableRow(new List<string> { "01.01.2022", "New Description1", "11 €", "21 €" }),
                new TableRow(new List<string> { "02.02.2022", "New Description2", "12 €", "22 €" }),
                new TableRow(new List<string> { "03.03.2022", "New Description3", "13 €", "23 €" }),
            };
            var table = new TableValues(tableCellNames, tableRows);

			var processor = new DocxProcessor();
			processor.CreateNewFileFromTemplate(docxTemplateFilename, docxOutputFilename, fields, table);
		}

		private void Button_EditDocx_Click(object sender, RoutedEventArgs e)
		{
			ExternalPrograms.OpenFileInEditor(Path.GetFullPath("GeneratedFile.docx"));
		}

		private void Button_GeneratePDF_Click(object sender, RoutedEventArgs e)
		{
			DocxToPdfConverter.Convert(docxOutputFilename, pdfOutputFilename, hide:false, exit:true);
		}

		private void Button_OpenPDF_Click(object sender, RoutedEventArgs e)
		{
			ExternalPrograms.OpenFileInEditor(Path.GetFullPath(pdfOutputFilename));
		}

		private void Button_SendPDF_Click(object sender, RoutedEventArgs e)
		{
			EmailParameters parameters = CreateEmailParameterSet();

			// Common version, using the mailto protocol. This isn't able to add an attachment.
			parameters.Body = parameters.Body.Replace("\n", "%0D");
			var executor = new ShellExecute();
			executor.Execute($"mailto:{parameters.To}?cc={parameters.Cc}&bcc={parameters.Bcc}&subject={parameters.Subject}&body={parameters.Body}");
		}

		private void Button_SendPDFOutlook_Click(object sender, RoutedEventArgs e)
		{
			EmailParameters parameters = CreateEmailParameterSet();

			// Version using the Microsoft Office Interop API
			OfficeInterop.OpenNewEmailWindow(parameters);
		}

		private static EmailParameters CreateEmailParameterSet()
		{
			return new EmailParameters()
			{
				Subject     = "My subject",
				Body        = "Dear Sirs,\n\nplease find attached my invoice.\n\nBest regards\nOliver",
				Sender      = "mail@oliver-abraham.de",
				To          = "john1@example.com",
				Cc          = "john2@example.com",
				Bcc         = "john3@example.com",
				Attachments = new List<string>() { Path.GetFullPath(pdfOutputFilename) }
			};
		}
	}
}
