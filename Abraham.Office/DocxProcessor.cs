using System.IO.Compression;
using System.Xml;

namespace Abraham.Office;

public class DocxProcessor
{
	#region ------------- Methods -------------------------------------------------------------
	public void CreateNewFileFromTemplate(string templateDocxFilename, string outputDocxFilename, List<FieldValue> fieldValues, TableValues tableValues)
    {
        if (!File.Exists(templateDocxFilename))
            throw new FileNotFoundException($"The Word document {templateDocxFilename} does not exist.");
		
		var subdirectory = "temp";
		if (Directory.Exists(subdirectory))
			Directory.Delete(subdirectory, recursive:true);

		ZipFile.ExtractToDirectory(templateDocxFilename, subdirectory);

		UpdateContents(subdirectory, fieldValues, tableValues);

		if (File.Exists(outputDocxFilename))
			File.Delete(outputDocxFilename);

		ZipFile.CreateFromDirectory(subdirectory, outputDocxFilename);
    }
	#endregion



	#region ------------- Implementation ------------------------------------------------------
    /// <summary>
    /// Update the document. Write a new value into the first cell on the first row of the first worksheet. 
    /// </summary>
    private static void UpdateContents(string subdirectory, List<FieldValue> fieldValues, TableValues tableValues)
	{
		FillParagraphs(Path.Combine(subdirectory, "word", "document.xml"), fieldValues);
		FillParagraphs(Path.Combine(subdirectory, "word", "header1.xml"), fieldValues);
		FillParagraphs(Path.Combine(subdirectory, "word", "header2.xml"), fieldValues);
		FillParagraphs(Path.Combine(subdirectory, "word", "header3.xml"), fieldValues);
		FillParagraphs(Path.Combine(subdirectory, "word", "footer1.xml"), fieldValues);
		FillParagraphs(Path.Combine(subdirectory, "word", "footer2.xml"), fieldValues);
		FillParagraphs(Path.Combine(subdirectory, "word", "footer3.xml"), fieldValues);

		FillTable(Path.Combine(subdirectory, "word", "document.xml"), tableValues).GetAwaiter().GetResult();
	}

	private static void FillParagraphs(string filename, List<FieldValue> fieldValues)
	{
		if (!File.Exists(filename))
			return;

		var contents = File.ReadAllText(filename);

		foreach (var field in fieldValues)
			contents = contents.Replace(field.Name, field.Value);

		File.WriteAllText(filename, contents);
	}

	private static async Task FillTable(string filename, TableValues tableValues)
	{
		var tableRowCount = tableValues.RowValues.Count;
		var currentRowCount = 0;
		var doc = new XmlDocument();
		doc.Load(filename);

		try
		{
			var body = doc.DocumentElement.FirstChild;
			var paragraphs = body.ChildNodes.Cast<XmlNode>().ToList();
			var tables = paragraphs.Where(x => x.Name == "w:tbl").ToList();
			var firstTable = tables.First();
			var rows = firstTable.ChildNodes.Cast<XmlNode>().ToList().Where(x => x.Name == "w:tr");
			foreach(var row in rows)
			{
				if (currentRowCount < tableRowCount)
					ProcessRow(tableValues, currentRowCount, row);
				else
					firstTable.RemoveChild(row);

				currentRowCount++;
			}
		}
		catch (Exception ex)
		{

		}

		doc.Save(filename);


		#region old
		//var contents = File.ReadAllText(filename);
		//var stream = new MemoryStream(Encoding.UTF8.GetBytes(contents));
		//var xPath = new XPathDocument(stream);
		//var navigator = xPath.CreateNavigator();
		//
		//XPathExpression query = navigator.Compile("body");
		//
		////Do some BS to get the default namespace to actually be called w. 
		//var nameSpace = new XmlNamespaceManager(navigator.NameTable);
		//nameSpace.AddNamespace("w", "JUST_ANY_NAMESPACE");
		//query.SetContext(nameSpace);
		//var singleNode = navigator.SelectSingleNode(query);
		//
		//// reading the list:
		//query = navigator.Compile("body/tbl/tr");
		//nameSpace = new XmlNamespaceManager(navigator.NameTable);
		//nameSpace.AddNamespace("w", "JUST_ANY_NAMESPACE");
		//query.SetContext(nameSpace);
		//var listOfNodes = navigator.SelectDescendants("w:p", "JUST_ANY_NAMESPACE", matchSelf:false);
		//var values = "";
		//foreach (var item in listOfNodes)
		//	values += $"{item},";






		//var config = Configuration.Default.WithXml();
		//var context = BrowsingContext.New(config);
		//var document = await context.OpenAsync(req => req.Content(filename));
		//
		//// find and fetch the table rows:
		//var body = document.GetElementsByTagName("body");
		//var text = body[0].ChildNodes;
		//var body2 = body.Where(x => x.TagName == "p").ToList();
		//var paragraphs = document.All.Where(x => x.TagName == "p").ToList();
		//
		//// find and fetch the list:
		////var listNodes = elements.Where(x => x.NodeName == "p").ToList();
		////var listItems = listNodes[0].GetElementsByTagName("MyListItem");
		////var listItemValues = listItems.Select(x => x.InnerHtml);
		////var listItemsAsString = string.Join(',', listItemValues);
		#endregion
	}

	private static void ProcessRow(TableValues tableValues, int currentRowCount, XmlNode row)
	{
		var cells = row.ChildNodes.Cast<XmlNode>().ToList().Where(x => x.Name == "w:tc");
		foreach (var cell in cells)
		{
			var cellParagraphs = cell.ChildNodes.Cast<XmlNode>().ToList().Where(x => x.Name == "w:p");
			foreach (var cellPara in cellParagraphs)
			{
				var runs = cellPara.ChildNodes.Cast<XmlNode>().ToList().Where(x => x.Name == "w:r");
				foreach (var run in runs)
				{
					var texts = run.ChildNodes.Cast<XmlNode>().ToList().Where(x => x.Name == "w:t");
					foreach (var text in texts)
					{
						for (int i = 0; i < tableValues.CellNames.Count; i++)
						{
							if (text.InnerText == tableValues.CellNames[i])
								text.InnerText = tableValues.RowValues[currentRowCount].CellValues[i];
						}
					}
				}
			}
		}
	}

	private static void RemoveRow(XmlDocument doc, XmlNode row)
	{
		row.RemoveChild(doc);
	}

	private static void ProcessParagraph(List<FieldValue> fieldValues, TableValues tableValues, object paragraphObj)
	{
		//else if (paragraphObj is NPOI.OpenXmlFormats.Wordprocessing.CT_Tbl)
		//{
		//	var table = paragraphObj as NPOI.OpenXmlFormats.Wordprocessing.CT_Tbl;
		//	var rowCounter = 0;
		//	foreach (var rowObj in table.Items1)
		//	{
		//		if (rowCounter + 1 > tableValues.RowValues.Count)
		//		{
		//			while (table.Items1.Count > tableValues.RowValues.Count)
		//				table.RemoveTr(rowCounter);
		//			break;
		//		}
		//		else
		//		{
		//			var row = rowObj as NPOI.OpenXmlFormats.Wordprocessing.CT_Row;
		//			if (row != null)
		//			{
		//				var columnCounter = 0;
		//				foreach (var cellObj in row.Items)
		//				{
		//					var cell = cellObj as NPOI.OpenXmlFormats.Wordprocessing.CT_Tc;
		//					if (cell != null && cell.Items != null)
		//					{
		//						foreach (var paragraphObj2 in cell.Items)
		//						{
		//							var paragraph = paragraphObj2 as NPOI.OpenXmlFormats.Wordprocessing.CT_P;
		//							foreach (var regionObj in paragraph.Items)
		//							{
		//								var region = regionObj as NPOI.OpenXmlFormats.Wordprocessing.CT_R;
		//								foreach (var textObj in region.Items)
		//								{
		//									var text = textObj as NPOI.OpenXmlFormats.Wordprocessing.CT_Text;
		//									var cellName = tableValues.CellNames[columnCounter];
		//									if (text.Value.Contains(cellName))
		//									{
		//										var newValue = tableValues.RowValues[rowCounter].CellValues[columnCounter];
		//										text.Value = text.Value.Replace(cellName, newValue);
		//									}
		//								}
		//							}
		//						}
		//					}
		//					columnCounter++;
		//				}
		//			}
		//		}
		//		rowCounter++;
		//	}
		//}
		//else
		//{
		//}
	}
	#endregion
}
