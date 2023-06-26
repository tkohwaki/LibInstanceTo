namespace LibInstanceTo.Excel;
using OfficeOpenXml;

public class InstanceToExcel<T,V> : ConvertBase<T,V> 
    where V : new() {
    public V Converter { get; set; }
    private ExcelPackage? pkg;
    private List<ConvertDef> defs;
    protected InstanceToExcel(string DefFile) {
        Converter = new V();
        using(FileStream stm = new FileStream(DefFile,FileMode.Open)) {
            defs = LoadDef(stm,OutputTypes.Excel);
        }
    }
    public InstanceToExcel(string ExcelFile,string DefFile)
        : this(DefFile) {
        pkg = new ExcelPackage(new FileInfo(ExcelFile));
    }
    public InstanceToExcel(Stream ExcelStream,string DefFile)
        : this(DefFile) {
        pkg = new ExcelPackage(ExcelStream);
    }
    public void ConvertOne(int Row,T Inst) {
        foreach(var itm in defs) {
            ExcelWorksheet sheet = pkg!.Workbook.Worksheets[itm.SheetName!];
            int col = itm.Index;
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
    public void Convert(List<T> Instances) {
        int row = StartRow;
        foreach(var itm in Instances) {
            ConvertOne(row,itm);
            row++;
        }
    }
    public void SaveAs(Stream OutStream) {
        pkg!.SaveAs(OutStream);
        OutStream.Flush();
    }
    public void SaveAs(string FileName) {
        pkg!.SaveAs(new FileInfo(FileName));
    }
}
