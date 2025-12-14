using SkiaSharp;

namespace WordTemplater
{
  internal static class Constant
  {
    public const string HTML_PATTERN = "<html><head><meta charset=\"UTF-8\"></head><body>{0}</body></html>";
    public const string PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    public const long PIXEL_PER_INCH = 914400L;
    public const long TWIP_PER_INCH = 1440L;
    public const double DEFAULT_DPI = 95.9865952;
    public const int DEFAULT_BARCODE_HIGHT = 70;
    public const int DEFAULT_BARCODE_BARWIDTH = 3;
    public const int DEFAULT_QRCODE_SIZE = 512;
    public const int DEFAULT_QRCODE_BORDER = 0;
    public static readonly SKColor DEFAULT_DARK_COLOR = SKColor.Parse("000000");
    public static readonly SKColor DEFAULT_LIGHT_COLOR = SKColor.Empty;
    public static readonly string DEFAULT_IMAGE_FILE_NAME = "{0}.png";
    public static readonly string PARSER_PARAM_REGEX = @"\s*(?:(?:""([^""]*(?:'[^""]*)*)"")|(?:'([^']*(?:\""[^']*)*)')|([^,'""]+))\s*(?:,|$)";
    public static readonly string MERGEFIELD = "MERGEFIELD  ";
    public static readonly string MERGEFORMAT = "  \\* MERGEFORMAT";
    public static readonly string CURRENT_NODE = ".";
    public static readonly string CURRENT_INDEX = "_index";
    public static readonly string IS_LAST = "_last";

  }

  internal static class FunctionName
  {
    internal const string Default = "";
    internal const string Sub = "sub";
    internal const string Left = "left";
    internal const string Right = "right";
    internal const string Trim = "trim";
    internal const string Upper = "upper";
    internal const string Lower = "lower";
    internal const string If = "if";
    internal const string Currency = "currency";
    internal const string Percentage = "percentage";
    internal const string Replace = "replace";
    internal const string BarCode = "barcode";
    internal const string QRCode = "qrcode";
    internal const string Image = "image";
    internal const string Mermaid = "mermaid";
    internal const string Html = "html";
    internal const string Word = "word";
    internal const string Loop = "loop";
    internal const string Table = "table";
    internal const string EndLoop = "endloop";
    internal const string EndTable = "endtable";
    internal const string EndIf = "endif";
  }

  internal static class OperatorName
  {
    internal const string Gt = ">";
    internal const string Lt = "<";
    internal const string Eq1 = "==";
    internal const string Eq2 = "=";
    internal const string Neq1 = "!=";
    internal const string Neq2 = "<>";
    internal const string Geq = ">=";
    internal const string Leq = "<=";
  }

  internal enum CompareValue
  {
    Eq = 1,
    Gt = 2,
    Lt = 4
  }
}
