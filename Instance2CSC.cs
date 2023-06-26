namespace LibInstanceTo.CSVl;
using System.Text;
using System.Numerics;

public class InstanceToCSV<T,V> : ConvertBase<T,V> , IDisposable
    where V : new() {
    public V Converter { get; set; }
    private List<ConvertDef> defs;
    private StreamWriter outfile = null!;
    protected InstanceToCSV(string DefFile) {
        Converter = new V();
        using(FileStream stm = new FileStream(DefFile,FileMode.Open)) {
            defs = LoadDef(stm,OutputTypes.Excel);
        }
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
    public InstanceToCSV(string CSVFile,string DefFile)
        : this(DefFile) {
        outfile = new StreamWriter(CSVFile,false,Encoding.GetEncoding("shift_jis"));
    }
    public InstanceToCSV(Stream CSVStream,string DefFile)
        : this(DefFile) {
        outfile = new StreamWriter(CSVStream,Encoding.GetEncoding("shift_jis"));
    }
    public void ConvertOne(int Row,T Inst) {
        string?[] items = new string[defs.Max(v=>v.Index)+1];
        foreach(var itm in defs) {
            object? val = typeof(T).GetProperty(itm.PropertyName)!.GetValue(Inst);
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
    public void Convert(List<T> Instances) {
        int row = 1;
        foreach(var itm in Instances) {
            ConvertOne(row,itm);
            row++;
        }
        outfile.Close();
    }
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
    public void Dispose() {
        if (outfile != null) {
            outfile.Dispose();
        }
    }
}
