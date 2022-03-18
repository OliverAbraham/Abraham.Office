namespace Abraham.Office;

public class TableRow
{
	public List<string> CellValues { get; set; }
	public TableRow(List<string> cellValues)
	{
         CellValues = cellValues;
	}
}
