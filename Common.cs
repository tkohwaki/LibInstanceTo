namespace LibInstanceTo;
using System.Reflection;
using System.Xml.Linq;

/// <summary>
/// InstanceToXXX Base Class
/// </summary>
/// <typeparam name="T"></typeparam>
/// <typeparam name="V"></typeparam>
public class ConvertBase<T,V>
    where V : new() {
    /// <summary>
    /// Instance Variable
    /// </summary>
    protected int StartRow = 1;
    /// <summary>
    /// Load Convert Definition File
    /// </summary>
    /// <param name="stm">Definition File Stream</param>
    /// <param name="kind">OutputKind.Excel or OutputKind.CSV</param>
    /// <returns>Convert Definition List</returns>
    protected List<ConvertDef> LoadDef(Stream stm,OutputTypes kind) {
        List<ConvertDef> defs = new List<ConvertDef>();
        // Load XML from Stream
        XDocument doc = XDocument.Load(stm,LoadOptions.None);
        if (doc.Root == null) {
            throw new InvalidDataException("XML形式が空です。");
        }
        if (doc.Root.Element("Column") == null) {
            throw new InvalidDataException("カラム定義が存在しません");
        }
        if (doc.Root.Attribute("StartRow") != null) {
            StartRow = Convert.ToInt32(doc.Root.Attribute("StartRow")!.Value);
        }
        // Set Column Definition
        foreach(var itm in doc.Root.Elements("Column")) {
            ConvertDef def = new ConvertDef();
            // Index
            string? c = itm.Attribute("Index")?.Value;
            if (c == null) {
                throw new InvalidDataException("必須属性(Index)が存在しません");
            }
            // Start Row(Excel Only)
            def.Index = Convert.ToInt32(c);
            // PropertyName in Class
            c = itm.Attribute("Property")?.Value;
            if (c == null) {
                throw new InvalidDataException("必須項目(Property)が存在しません");
            }
            // Format
            c = itm.Attribute("Format")?.Value;
            def.Format = c;
            // Convert Function
            c = itm.Attribute("Func")?.Value;
            if (c == null) {
                def.Converter = null;
            } else {
                MethodInfo? m = typeof(V).GetMethod(c);
                if (m == null) {
                    throw new InvalidDataException("変換メソッドが指定クラスに存在しません");
                } else {
                    def.Converter = m;
                }
            }
            // Sheet(Excel Only)
            c = itm.Attribute("Sheet")?.Value;
            if (kind == OutputTypes.Excel && c == null) {
                throw new InvalidDataException("Excel出力には(Sheet)属性が必須です。");
            }
            def.SheetName = c;
            defs.Add(def);
        }
        return defs;
    }
}
/// <summary>
/// Convert Definition
/// </summary>
public class ConvertDef {
    public string PropertyName { get; set; } = null!;   // マッピングするプロパティ名
    public int Index { get; set; }  // カラム(Excelは1ベース,CSVは0ベース)
    public string? SheetName { get; set; }   // シート名(Excelのみ)
    public string? Format { get; set; } // 変換フォーマット(オプション)
    public MethodInfo? Converter { get; set; }  // 変換関数(オプション)
}
/// <summary>
/// OutputTypes Enum
/// </summary>
public enum OutputTypes {
    Excel,CSV
}