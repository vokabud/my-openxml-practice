namespace OpenXmlPractice.Core.Exceptions;

public class SheetNotFoundException : Exception
{
    public SheetNotFoundException() : base("Sheet not found.")
    {
    }

    public SheetNotFoundException(string name) : base($"Sheet {name} not found.")
    {
    }

    public SheetNotFoundException(string name, Exception innerException) : base($"Sheet {name} not found.", innerException)
    {
    }
}
