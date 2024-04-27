using System.Diagnostics;
using System.Globalization;
using Bogus;
using ExcelSerializer;
using ExcelSerializerConsole;
using static Bogus.DataSets.Name;

var sw = Stopwatch.StartNew();

Randomizer.Seed = new Random(8675309);

var fruit = new[] { "apple", "banana", "orange", "strawberry", "kiwi" };

var orderIds = 0;
var testOrders = new Faker<Order>()
    .StrictMode(true)
    .RuleFor(o => o.OrderId, f => orderIds++)
    .RuleFor(o => o.Item, f => f.PickRandom(fruit))
    .RuleFor(o => o.Quantity, f => f.Random.Number(-10, 10))
    .RuleFor(o => o.LotNumber, f => f.Random.Int(0, 100).OrNull(f, .8f));

var userIds = 0;
var testUsers = new Faker<User>()
    .CustomInstantiator(f => new User(userIds++, f.Random.Replace("###-##-####")))
    .RuleFor(u => u.Gender, f => f.PickRandom<Gender>())
    .RuleFor(u => u.FirstName, (f, u) => f.Name.FirstName(u.Gender))
    .RuleFor(u => u.LastName, (f, u) => f.Name.LastName(u.Gender))
    .RuleFor(u => u.Avatar, f => f.Internet.Avatar())
    .RuleFor(u => u.UserName, (f, u) => f.Internet.UserName(u.FirstName, u.LastName))
    .RuleFor(u => u.Email, (f, u) => f.Internet.Email(u.FirstName, u.LastName))
    .RuleFor(u => u.SomethingUnique, f => $"Value {f.UniqueIndex}")
    .RuleFor(u => u.TimeStamp, f => f.Date.Recent())
    .RuleFor(u => u.CreateTime, f => f.Date.Recent())
    .RuleFor(u => u.DateOnlyValue, f => f.Date.RecentDateOnly())
    .RuleFor(u => u.TimeOnlyValue, f => f.Date.RecentTimeOnly())
    .RuleFor(u => u.TimeSpanValue, f => f.Date.Recent() - f.Date.Past())
    .RuleFor(u => u.DateTimeOffsetValue, f => f.Date.Recent())
    .RuleFor(u => u.Fallback, (f, u) => (object)userIds)
    .RuleFor(u => u.Uri, f => new Uri(f.Internet.Url()))
    .RuleFor(u => u.SomeGuid, f => f.Random.Guid())
    .RuleFor(u => u.SendFlag, f => userIds % 3 == 0)
    .RuleFor(u => u.CartId, f => f.Random.Guid())
    .RuleFor(u => u.FullName, (f, u) => u.FirstName + " " + u.LastName)
    .RuleFor(u => u.Orders, f => testOrders.Generate(3).ToList())
    .RuleFor(o => o.DoubleValue, f => f.Random.Double(-1000, 1000))
    .RuleFor(o => o.Char, f => (char)f.Random.Int(65, 65 + 26))
    .RuleFor(o => o.Escape, f => "</>\"'&");

var Users = testUsers.Generate(10);

sw.Stop();
Console.WriteLine($"testUsers.Generate count:{Users.Count:#,##0} duration:{sw.ElapsedMilliseconds:#,##0}ms");
sw.Restart();

var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.CurrentCulture,
    MaxDepth = 32,
    Provider = ExcelSerializerProvider.Create(
        [ExcelSerializerProvider.Default]),
    HasHeaderRecord = true,
    AutoFitColumns = true,
    AutoFilter = true,
};

var fileName = Path.Combine(Environment.CurrentDirectory, "test.xlsx");
if (File.Exists(fileName))
    File.Delete(fileName);

ExcelSerializer.ExcelSerializer.ToFile(Users, fileName, newConfig);

sw.Stop();

Console.WriteLine($"ExcelSerializer.ToFile duration:{sw.ElapsedMilliseconds:#,##0}ms");
Console.WriteLine($"Excel file created. Please check the file. {fileName}");

Console.WriteLine();
Console.WriteLine("press any key...");

Console.ReadLine();
