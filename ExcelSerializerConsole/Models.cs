using ExcelSerializerLib;

namespace ExcelSerializerConsole;

public class Order
{
    public int OrderId { get; set; }
    public string Item { get; set; } = "";
    public int Quantity { get; set; }
    public int? LotNumber { get; set; }
}

public class User(int userId, string ssn)
{
    public int Id { get; set; } = userId;
    public string FirstName { get; set; } = "";
    public string LastName { get; set; } = "";
    public string FullName { get; set; } = "";
    public string UserName { get; set; } = "";
    public string Email { get; set; } = "";
    public string SomethingUnique { get; set; } = "";
    public Guid SomeGuid { get; set; }
    public bool SendFlag { get; set; }

    public string Avatar { get; set; } = "";
    public Guid CartId { get; set; }
    public string SSN { get; set; } = ssn;
    [ExcelSerializer(typeof(UnixSecondsSerializer))]
    public DateTime TimeStamp { get; set; }
    public DateTime CreateTime { get; set; }
    public DateOnly DateOnlyValue { get; set; }
    public TimeOnly TimeOnlyValue { get; set; }
    public TimeSpan TimeSpanValue { get; set; }
    public DateTimeOffset DateTimeOffsetValue { get; set; }
    public object? Fallback { get; set; }
    public Uri? Uri { get; set; }
    public Bogus.DataSets.Name.Gender Gender { get; set; }

    public List<Order> Orders { get; set; } = [];
    public double DoubleValue { get; set; }
    public char Char { get; set; }
    public string Escape { get; set; } = "";
}
