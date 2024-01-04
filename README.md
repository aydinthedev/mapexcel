# ExcelMap

ExcelMap is a lightweight and easy-to-use library for mapping Excel files to C# objects. This library simplifies the process of reading from and writing to Excel files in your C# applications.

## Quick Start

### Usage

#### Defining a Model

```csharp
public class Foo
{
    public string Name { get; set; }

    public DateTime Created { get; set; }
}

// Call this method in your application's entry point
// E.g. Program.cs
ExcelTypeRegistry.Register<Foo>(foo =>
{
    foo
        .Property(p => p.Name, prop => { prop.Column("A"); })
        .Property(p => p.Created, prop => { prop.Column("B"); });
});
```

#### Writing to Excel

```csharp
using (var workbook = new ExcelWorkbook())
{
    workbook.Write(new Foo
    {
        Name = "bar",
        Created = DateTime.Now
    });
    workbook.SaveAs("HelloWorld.xlsx");
}
```

#### Reading from Excel

```csharp
using (var workbook = new ExcelWorkbook("HelloWorld.xlsx"))
{
    foreach (var rowData in workbook.Read<Foo>())
    {
        Console.WriteLine(rowData.RowNumber);

        var foo = rowData.Entity;
        Console.WriteLine(foo.Name);
        Console.WriteLine(foo.Created);
    }
}
```