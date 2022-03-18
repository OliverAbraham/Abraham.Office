namespace Abraham.Office;

public class FieldValue
{
	public string Name { get; set; }
	public string Value { get; set; }
	public FieldValue(string name, string value)
	{
        Name = name;
        Value = value;
	}
}
