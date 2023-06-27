namespace LibInstanceTo.Excel;
using OfficeOpenXml;

/// <summary>
/// Class Instance to Excel Column Class
/// </summary>
/// <typeparam name="T">Instance Type</typeparam>
/// <typeparam name="V">Converter Type</typeparam>
public class InstanceToExcel<T,V> : ConvertBase<T,V>, IDisposable
    where V : new() {
    /// <summary>
    /// Converter Instance
    /// </summary>
    /// <value></value>
    public V Converter { get; set; }
    /// <summary>
    /// Excel Package
    /// </summary>
    private ExcelPackage? pkg;
    /// <summary>
    /// Convert Definition List
    /// </summary>
    private List<ConvertDef> defs;
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="DefFile">Definition File Name</param>
    protected InstanceToExcel(string DefFile) {
        // Create Converter Instance
        Converter = new V();
        // Load Definition List
        using(FileStream stm = new FileStream(DefFile,FileMode.Open)) {
            defs = LoadDef(stm,OutputTypes.Excel);
        }
    }
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="ExcelFile">Input/Output Excel File Name</param>
    /// <param name="DefFile">Definition File Name</param>
    public InstanceToExcel(string? ExcelFile,string DefFile)
        : this(DefFile) {
        if (!string.IsNullOrEmpty(ExcelFile)) {
            pkg = new ExcelPackage(new FileInfo(ExcelFile));
        } else {
            pkg = new ExcelPackage();
        }
    }
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="ExcelStream">Input/Output Excel Stream</param>
    /// <param name="DefFile">Definition File Name</param>
    public InstanceToExcel(Stream ExcelStream,string DefFile)
        : this(DefFile) {
        if (ExcelStream != null) {
            pkg = new ExcelPackage(ExcelStream);
        } else {
            pkg = new ExcelPackage();
        }
    }
    /// <summary>
    /// Convert Single Instance
    /// </summary>
    /// <param name="Row">Row</param>
    /// <param name="Inst">Instance</param>
    public void ConvertOne(int Row,T Inst) {
        // Each Definition
        foreach(var itm in defs) {
            // Get Worksheet
            ExcelWorksheet sheet = pkg!.Workbook.Worksheets[itm.SheetName!];
            if (sheet == null) {
                pkg!.Workbook.Worksheets.Add(itm.SheetName!);
                sheet = pkg!.Workbook.Worksheets[itm.SheetName!];
            }
            // Excel Column
            int col = itm.Index;
            // Get Value from Instance.PropertyName
            object? val = typeof(T).GetProperty(itm.PropertyName)!.GetValue(Inst);
            if (itm.Converter == null) {
                sheet.Cells[Row,col].Value = val;
            } else {
                sheet.Cells[Row,col].Value = itm.Converter!.Invoke(Converter,new object?[] {val});
            }
            if (itm.Format != null) {
                sheet.Cells[Row,col].Style.Numberformat.Format = itm.Format;
            }
        }
    }
    /// <summary>
    /// Convert All
    /// </summary>
    /// <param name="Instances">Instance List</param>
    public void Convert(IEnumerable<T> Instances) {
        int row = StartRow;
        foreach(var itm in Instances) {
            ConvertOne(row,itm);
            row++;
        }
    }
    /// <summary>
    /// SaveAs(Stream)
    /// </summary>
    /// <param name="OutStream">Output Stream</param>
    public void SaveAs(Stream OutStream) {
        pkg!.SaveAs(OutStream);
        OutStream.Flush();
    }
    /// <summary>
    /// SaveAs(File)
    /// </summary>
    /// <param name="FileName">Output File Name</param>
    public void SaveAs(string FileName) {
        pkg!.SaveAs(new FileInfo(FileName));
    }
    /// <summary>
    /// Destructor
    /// </summary>
    public void Dispose() {
        if (pkg != null) {
            pkg.Dispose();
        }
    }
}
