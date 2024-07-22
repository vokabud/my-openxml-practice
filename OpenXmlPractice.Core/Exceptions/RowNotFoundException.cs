namespace OpenXmlPractice.Core.Exceptions;

public class RowNotFoundException : Exception
{
    public RowNotFoundException()
    {
    }

    public RowNotFoundException(string message) : base(message)
    {
    }

    public RowNotFoundException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
