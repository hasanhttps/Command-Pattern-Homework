using System.Data;
using iTextSharp.text;
using ClosedXML.Excel;
using iTextSharp.text.pdf;
using System.IO.Compression;


namespace Command_Pattern_Excel_Homework;

public class Product {
    public int Id { get; set; }
    public string Name { get; set; }
    public decimal Price { get; set; }
    public int Stock { get; set; }

    public override string ToString() {
        return $"Id : {Id}\nName : {Name}\nPrice : {Price}\nStock : {Stock}\n\n";
    }
}

// Receiver
public class ExcelFile<T> {
    private readonly List<T> _list;
    public string FileName => $"{typeof(T).Name}.xlsx";


    public ExcelFile(List<T> list) {
        _list = list;
    }

    public MemoryStream Create() {
        var wb = new XLWorkbook();
        var ds = new DataSet();

        ds.Tables.Add(GetTable());
        wb.Worksheets.Add(ds);

        var excelMemory = new MemoryStream();
        wb.SaveAs(excelMemory);

        return excelMemory;
    }


    private DataTable GetTable() {
        var table = new DataTable();
        
        var type = typeof(T);

        type.GetProperties()
            .ToList()
            .ForEach(x => table.Columns.Add(x.Name, x.PropertyType));


        _list.ForEach(x => {
            var values = type.GetProperties()
                                .Select(properyInfo => properyInfo
                                .GetValue(x, null))
                                .ToArray();

            table.Rows.Add(values);
        });

        return table;
    }
}

// Homework
public class PdfFile<T> {
    private readonly List<T> _list;
    public string FileName => $"{typeof(T).Name}.pdf";

    public PdfFile(List<T> list) {
        _list = list;
    }

    public MemoryStream Create() {
        Document document = new Document();
        MemoryStream memoryStream = new MemoryStream();
        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
        document.Open();

        // Create a list
        List list = new List(List.UNORDERED);
        foreach(var item in _list) {
            list.Add(item.ToString());
        } 

        // Add the list to the document
        document.Add(list);

        // Close the document
        document.Close();

        return memoryStream;
    }
}

public interface ITableActionCommand {
    void Execute();
    void ExecuteInZip();
}


public class CreateExcelTableActionCommand<T> : ITableActionCommand {
    private readonly ExcelFile<T> _excelFile;

    public CreateExcelTableActionCommand(ExcelFile<T> excelFile)
        => _excelFile = excelFile;


    public void Execute() {
        MemoryStream excelMemoryStream = _excelFile.Create();
        File.WriteAllBytes(_excelFile.FileName, excelMemoryStream.ToArray());
    }

    public void ExecuteInZip() {

        using (MemoryStream zipMemoryStream = _excelFile.Create()) {
            using (ZipArchive archive = new ZipArchive(zipMemoryStream, ZipArchiveMode.Create, true)) {
                var memoryStream = zipMemoryStream;
                string entryName = _excelFile.FileName;
                ZipArchiveEntry entry = archive.CreateEntry(entryName);
                using (Stream entryStream = entry.Open()) {
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    memoryStream.CopyTo(entryStream);
                }
            }
            File.WriteAllBytes("files.zip", zipMemoryStream.ToArray());
        }
    }
}

public class CreatePdfTableActionCommand<T> : ITableActionCommand {
    private readonly PdfFile<T> _pdfFile;

    public CreatePdfTableActionCommand(PdfFile<T> pdfFile)
        => _pdfFile = pdfFile;


    public void Execute() {
        MemoryStream pdfMemoryStream = _pdfFile.Create();
        File.WriteAllBytes(_pdfFile.FileName, pdfMemoryStream.ToArray());
    }

    public void ExecuteInZip() {

        using (MemoryStream zipMemoryStream = _pdfFile.Create()) {
            using (ZipArchive archive = new ZipArchive(zipMemoryStream, ZipArchiveMode.Read, true)) {
                var memoryStream = zipMemoryStream;
                string entryName = _pdfFile.FileName;
                ZipArchiveEntry entry = archive.CreateEntry(entryName);
                using (Stream entryStream = entry.Open()) {
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    memoryStream.CopyTo(entryStream);
                }
            }
            File.WriteAllBytes("files.zip", zipMemoryStream.ToArray());
        }
    }
}

// Invoker
class FileCreateInvoker {
    private ITableActionCommand _tableActionCommand;
    private List<ITableActionCommand> tableActionCommands = new List<ITableActionCommand>();

    public void SetCommand(ITableActionCommand tableActionCommand) {
        _tableActionCommand = tableActionCommand;
    }

    public void AddCommand(ITableActionCommand tableActionCommand) {
        tableActionCommands.Add(tableActionCommand);
    }

    public void CreateFile() {
        _tableActionCommand.Execute();
    }

    public void CreateFiles() {
        foreach (var tableActionCommand in tableActionCommands) {
            tableActionCommand.ExecuteInZip();
        }
    }
}

class Program {
    static void Main() {

        var products = Enumerable.Range(1, 30).Select(index =>
            new Product {
                Id = index,
                Name = $"Product {index}",
                Price = index + 100,
                Stock = index
            }
        ).ToList();

        ExcelFile<Product> receiver = new(products);
        PdfFile<Product> receiver2 = new(products);

        FileCreateInvoker invoker = new();
        //invoker.AddCommand(new CreateExcelTableActionCommand<Product>(receiver));
        invoker.AddCommand(new CreatePdfTableActionCommand<Product>(receiver2));
        invoker.CreateFiles();
    }
}