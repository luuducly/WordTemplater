using SkiaSharp;

namespace WordTemplater
{
    /// <summary>
    /// Contains constant values used throughout the WordTemplater library.
    /// </summary>
    internal static class Constant
    {
        /// <summary>
        /// HTML template pattern for rendering HTML content.
        /// </summary>
        public const string HTML_PATTERN =
            "<html><head><meta charset=\"UTF-8\"></head><body>{0}</body></html>";

        /// <summary>
        /// XML namespace for picture elements in OpenXML documents.
        /// </summary>
        public const string PICTURE_NAMESPACE =
            "http://schemas.openxmlformats.org/drawingml/2006/picture";

        /// <summary>
        /// Number of pixels per inch in Word documents.
        /// </summary>
        public const long PIXEL_PER_INCH = 914400L;

        /// <summary>
        /// Default DPI (Dots Per Inch) value used for image processing.
        /// </summary>
        public const double DEFAULT_DPI = 95.9865952;

        /// <summary>
        /// Default height for barcode images in pixels.
        /// </summary>
        public const int DEFAULT_BARCODE_HIGHT = 70;

        /// <summary>
        /// Default bar width for barcode images in pixels.
        /// </summary>
        public const int DEFAULT_BARCODE_BARWIDTH = 3;

        /// <summary>
        /// Default size for QR code images in pixels.
        /// </summary>
        public const int DEFAULT_QRCODE_SIZE = 512;

        /// <summary>
        /// Default border size for QR code images in pixels.
        /// </summary>
        public const int DEFAULT_QRCODE_BORDER = 0;

        /// <summary>
        /// Default dark color used for barcodes and QR codes.
        /// </summary>
        public static readonly SKColor DEFAULT_DARK_COLOR = SKColor.Parse("000000");

        /// <summary>
        /// Default light color used for barcodes and QR codes.
        /// </summary>
        public static readonly SKColor DEFAULT_LIGHT_COLOR = SKColor.Empty;

        /// <summary>
        /// Default file name pattern for generated images.
        /// </summary>
        public static readonly string DEFAULT_IMAGE_FILE_NAME = "{0}.png";

        /// <summary>
        /// Regular expression pattern for parsing template parameters.
        /// </summary>
        public static readonly string PARSER_PARAM_REGEX =
            @"\s*(?:(?:""([^""]*(?:'[^""]*)*)"")|(?:'([^']*(?:\""[^']*)*)')|([^,'""]+))\s*(?:,|$)";

        /// <summary>
        /// Merge field marker used in Word templates.
        /// </summary>
        public static readonly string MERGEFIELD = "MERGEFIELD  ";

        /// <summary>
        /// Merge format marker used in Word templates.
        /// </summary>
        public static readonly string MERGEFORMAT = "  \\* MERGEFORMAT";

        /// <summary>
        /// Marker for current node in template expressions.
        /// </summary>
        public static readonly string CURRENT_NODE = ".";

        /// <summary>
        /// Marker for current index in template expressions.
        /// </summary>
        public static readonly string CURRENT_INDEX = "_index";

        /// <summary>
        /// Marker for last item in template expressions.
        /// </summary>
        public static readonly string IS_LAST = "_last";
    }

    /// <summary>
    /// Contains function names used in template expressions.
    /// </summary>
    internal static class FunctionName
    {
        /// <summary>
        /// Default function name (empty string).
        /// </summary>
        internal const string Default = "";

        /// <summary>
        /// Substring function name.
        /// </summary>
        internal const string Sub = "sub";

        /// <summary>
        /// Left substring function name.
        /// </summary>
        internal const string Left = "left";

        /// <summary>
        /// Right substring function name.
        /// </summary>
        internal const string Right = "right";

        /// <summary>
        /// Trim function name.
        /// </summary>
        internal const string Trim = "trim";

        /// <summary>
        /// Convert to uppercase function name.
        /// </summary>
        internal const string Upper = "upper";

        /// <summary>
        /// Convert to lowercase function name.
        /// </summary>
        internal const string Lower = "lower";

        /// <summary>
        /// Conditional if function name.
        /// </summary>
        internal const string If = "if";

        /// <summary>
        /// Currency formatting function name.
        /// </summary>
        internal const string Currency = "currency";

        /// <summary>
        /// Percentage formatting function name.
        /// </summary>
        internal const string Percentage = "percentage";

        /// <summary>
        /// String replace function name.
        /// </summary>
        internal const string Replace = "replace";

        /// <summary>
        /// Barcode generation function name.
        /// </summary>
        internal const string BarCode = "barcode";

        /// <summary>
        /// QR code generation function name.
        /// </summary>
        internal const string QRCode = "qrcode";

        /// <summary>
        /// Image insertion function name.
        /// </summary>
        internal const string Image = "image";

        /// <summary>
        /// HTML content insertion function name.
        /// </summary>
        internal const string Html = "html";

        /// <summary>
        /// Word document insertion function name.
        /// </summary>
        internal const string Word = "word";

        /// <summary>
        /// Loop start function name.
        /// </summary>
        internal const string Loop = "loop";

        /// <summary>
        /// Table start function name.
        /// </summary>
        internal const string Table = "table";

        /// <summary>
        /// Loop end function name.
        /// </summary>
        internal const string EndLoop = "endloop";

        /// <summary>
        /// Table end function name.
        /// </summary>
        internal const string EndTable = "endtable";

        /// <summary>
        /// Conditional if end function name.
        /// </summary>
        internal const string EndIf = "endif";
    }

    /// <summary>
    /// Contains operator names used in template expressions.
    /// </summary>
    internal static class OperatorName
    {
        /// <summary>
        /// Greater than operator.
        /// </summary>
        internal const string Gt = ">";

        /// <summary>
        /// Less than operator.
        /// </summary>
        internal const string Lt = "<";

        /// <summary>
        /// Equal operator (==).
        /// </summary>
        internal const string Eq1 = "==";

        /// <summary>
        /// Equal operator (=).
        /// </summary>
        internal const string Eq2 = "=";

        /// <summary>
        /// Not equal operator (!=).
        /// </summary>
        internal const string Neq1 = "!=";

        /// <summary>
        /// Not equal operator (<>).
        /// </summary>
        internal const string Neq2 = "<>";

        /// <summary>
        /// Greater than or equal operator.
        /// </summary>
        internal const string Geq = ">=";

        /// <summary>
        /// Less than or equal operator.
        /// </summary>
        internal const string Leq = "<=";
    }

    /// <summary>
    /// Represents comparison values used in template expressions.
    /// </summary>
    internal enum CompareValue
    {
        /// <summary>
        /// Equal comparison.
        /// </summary>
        Eq = 1,

        /// <summary>
        /// Greater than comparison.
        /// </summary>
        Gt = 2,

        /// <summary>
        /// Less than comparison.
        /// </summary>
        Lt = 4
    }
}
