namespace LibInstanceTo.CSV;
using System.Text;
using System.Numerics;

/// <summary>
/// Class Instance to CSV File Class
/// </summary>
/// <typeparam name="T">Instance Type</typeparam>
/// <typeparam name="V">Converter Type</typeparam>
public class InstanceToCSV<T,V> : ConvertBase<T,V> , IDisposable
    where V : new() {
    /// <summary>
    /// Converter Class Instance
    /// </summary>
    /// <value></value>
    public V Converter { get; set; }
    /// <summary>
    /// Convert Definition List
    /// </summary>
    private List<ConvertDef> defs;
    /// <summary>
    /// Output CSV File Stream
    /// </summary>
    private StreamWriter outfile = null!;
    private int CurrentRow = 0;
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="DefFile">Convert Definition File</param>
    protected InstanceToCSV(string DefFile) {
        // Create Converter Instance
        Converter = new V();
        // Load Definition
        using(FileStream stm = new FileStream(DefFile,FileMode.Open)) {
            defs = LoadDef(stm,OutputTypes.Excel);
        }
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
    /// <summary>
    /// Constructor(File)
    /// </summary>
    /// <param name="CSVFile">Output CSV File Name</param>
    /// <param name="DefFile">Convert Definition File Name</param>
    public InstanceToCSV(string CSVFile,string DefFile)
        : this(DefFile) {
        outfile = new StreamWriter(CSVFile,false,Encoding.GetEncoding("shift_jis"));
    }
    /// <summary>
    /// Constructor(Stream)
    /// </summary>
    /// <param name="CSVStream">Output CSV Stream</param>
    /// <param name="DefFile">Convert Definition File Name</param>
    public InstanceToCSV(Stream CSVStream,string DefFile)
        : this(DefFile) {
        outfile = new StreamWriter(CSVStream,Encoding.GetEncoding("shift_jis"));
    }
    /// <summary>
    /// Convert Single Instance
    /// </summary>
    /// <param name="Row"></param>
    /// <param name="Inst"></param>
    public void ConvertOne(T Inst) {
        string?[] items = new string[defs.Max(v=>v.Index)+1];
        foreach(var itm in defs) {
            object? val = itm.Property!.GetValue(Inst);
            if (itm.Converter == null) {
                if (itm.Format == null) {
                    items[itm.Index] = $"{val}";
                } else {
                    items[itm.Index] = GetFormattedValue(val,itm.Format);
                }
            } else {
                object? v = itm.Converter!.Invoke(Converter,new object?[] {val});
                if (itm.Format != null) {
                    items[itm.Index] = GetFormattedValue(v,itm.Format);
                } else {
                    items[itm.Index] = $"{v}";
                }
            }
        }
        string outstr = "";
        for(int i=0; i < items.Length; i++) {
            outstr += $"{items[i]},";
        }
        if (!string.IsNullOrEmpty(outstr)) {
            outstr = outstr.Substring(0,outstr.Length-1);
        }
        outfile.WriteLine(outstr);
    }
    /// <summary>
    /// Convert Instance List
    /// </summary>
    /// <param name="Instances"></param>
    public void Convert(List<T> Instances) {
        if (CurrentRow == 0) {
            CurrentRow = StartRow;
        }
        foreach(var itm in Instances) {
            ConvertOne(itm);
            CurrentRow++;
        }
    }
    /// <summary>
    /// Close Output CSV File
    /// </summary>
    public void Close() {
        outfile.Close();
    }
    /// <summary>
    /// Get Formatted Value
    /// </summary>
    /// <param name="val">Value</param>
    /// <param name="format">Format</param>
    /// <returns>Formatted String</returns>
    private string? GetFormattedValue(object? val, string format) {
        Type? t = val?.GetType();
        if (t != null) {
            if (t.GetInterface("System.IFormattable") != null) {
                var v = (IFormattable)val!;
                return v.ToString(format,null);
            } else {
                return $"{val}";
            }
        } else {
            return null;
        }
    }
    /// <summary>
    /// Destructor
    /// </summary>
    public void Dispose() {
        if (outfile != null) {
            outfile.Dispose();
        }
    }
}
