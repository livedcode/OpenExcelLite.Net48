# OpenExcelLite.Net48

A lightweight, schema-safe Excel (XLSX) generator for **.NET Framework 4.8 / 4.8.1**  
built on the official **OpenXML 3.3.0** SDK — with **no Excel interop**, **no COM**, and **zero dependencies**.

This library mirrors the simplicity and API design philosophy of **OpenExcelLite (modern .NET)**  
while remaining fully compatible with legacy .NET Framework applications.

---

## ✨ Features

- Create Excel files entirely **in memory** or **save to file**
- **Multi-sheet** workbook support
- **Schema-safe** rows (column count validated)
- **AddRow()**, **AddRows()**, **AddEmptyRows()**
- **HyperlinkCell** for clickable links
- **Auto column width** calculation
- **Header styling** (bold + background)
- **Freeze panes** (top rows / left columns)
- **Date** handling using OADate numeric format
- Works on **.NET Framework 4.8 / 4.8.1**
- Zero external dependencies (other than OpenXML SDK)

---

## 📦 Installation

### Project Reference

```xml
<ItemGroup>
  <ProjectReference Include="OpenExcelLite.Net48.csproj" />
</ItemGroup>


```xml
<ItemGroup>
  <ProjectReference Include="OpenExcelLite.Net48.csproj" />
</ItemGroup>
```

Or use NuGet (if published):

```
Install-Package OpenExcelLite.Net48
```

---

## 🚀 Usage

### ✔ Simple Example

```csharp
var workbook = new ExcelWorkbook();
var sheet = workbook.AddSheet("Users");

sheet.AddRow("Id", "Name", "Email");
sheet.AddRow(1, "Alex", "alex@test.com");
sheet.AddRow(2, "Bella", "bella@test.com");

workbook.SaveToFile("users.xlsx");
```

---

### ✔ Hyperlink Example

```csharp
var wb = new ExcelWorkbook();
var sheet = wb.AddSheet("Links");

sheet.AddRow("Title", "URL");
sheet.AddRow("OpenAI", new HyperlinkCell("Visit", "https://openai.com"));

wb.SaveToFile("links.xlsx");
```

---

### ✔ Multi-Sheet Example

```csharp
var wb = new ExcelWorkbook();

var products = wb.AddSheet("Products");
products.AddRow("Id", "Name", "Price");
products.AddRow(1, "Keyboard", 129.90m);

var orders = wb.AddSheet("Orders");
orders.AddRow("OrderId", "Total");
orders.AddRow(1001, 199.80m);

wb.SaveToFile("multi_sheet.xlsx");
```

---

### ✔ In-Memory Array Example

```csharp
var wb = new ExcelWorkbook();
wb.AddSheet("Data").AddRow("A", "B", "C");

byte[] bytes = wb.ToArray();
File.WriteAllBytes("array.xlsx", bytes);
```

---

## 📁 Samples

See the `samples/OpenExcelLite.Net48.Sample` folder for complete demo cases:

- In-memory examples  
- Hyperlinks  
- Empty rows (before or after header)  
- Multi-sheet  
- Multi-sheet hyperlinks  
- 10-sheet generation  

---

## 🧪 Unit Tests

The `tests/OpenExcelLite.Net48.Tests` project includes coverage for:

- Multi-sheet workbooks
- Hyperlink relationships
- Auto column widths
- Date cells (OADate)
- Freeze panes correctness

---

## 🔧 Compatibility

| Feature | Supported |
|--------|-----------|
| .NET Framework 4.8 / 4.8.1 | ✔ |
| C# 7.3 | ✔ |
| OpenXML 3.3.0 | ✔ |
| Streaming APIs (Span/ArrayPool) | ✖ Not supported on .NET48 |
| Async APIs | ✖ |

---

## 📄 License

This project is released under the **MIT License**.

See: `LICENSE`

