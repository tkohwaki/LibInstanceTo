namespace LibInstanceTo;
using System.Reflection;
using System.Xml.Linq;

public class ConvertBase<T,V>
    where V : new() {
    protected int StartRow = 1;
    public List<ConvertDef> LoadDef(Stream stm,OutputTypes kind) {
        List<ConvertDef> defs = new List<ConvertDef>();
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
        foreach(var itm in doc.Root.Elements("Column")) {
            ConvertDef def = new ConvertDef();
            string? c = itm.Attribute("Index")?.Value;
            if (c == null) {
                throw new InvalidDataException("必須属性(Index)が存在しません");
            }
            def.Index = Convert.ToInt32(c);
            c = itm.Attribute("Property")?.Value;
            if (c == null) {
                throw new InvalidDataException("必須項目(Property)が存在しません");
            } 
            c = itm.Attribute("Format")?.Value;
            def.Format = c;
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
public class ConvertDef {
    public string PropertyName { get; set; } = null!;   // マッピングするプロパティ名
    public int Index { get; set; }  // カラム(Excelは1ベース,CSVは0ベース)
    public string? SheetName { get; set; }   // シート名(Excelのみ)
    public string? Format { get; set; } // 変換フォーマット(オプション)
    public MethodInfo? Converter { get; set; }  // 変換関数(オプション)
}
public enum OutputTypes {
    Excel,CSV
}