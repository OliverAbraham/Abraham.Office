namespace Abraham.Office;

public class TableValues
{
	public List<string> CellNames { get; set; }
	public List<TableRow> RowValues { get; set; }
	public TableValues(List<string> cellNames, List<TableRow> rowValues)
	{
        CellNames = cellNames;
        RowValues = rowValues;
	}
}
