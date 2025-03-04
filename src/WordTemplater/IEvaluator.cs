using System.Text.RegularExpressions;

namespace WordTemplater
{
    /// <summary>
    /// Defines the interface for evaluating and formatting template expressions.
    /// </summary>
    public interface IEvaluator
    {
        /// <summary>
        /// Evaluates and formats a data field value according to the specified parameters.
        /// </summary>
        /// <param name="fieldValue">The data field value to evaluate.</param>
        /// <param name="parameters">Additional parameters declared in the template file.</param>
        /// <returns>The formatted value to be displayed in the exported Word file.</returns>
        public string Evaluate(object fieldValue, List<object> parameters);
    }

    /// <summary>
    /// Default evaluator that applies basic string formatting.
    /// </summary>
    internal class DefaultEvaluator : IEvaluator
    {
        /// <summary>
        /// Evaluates the field value using the specified format string.
        /// </summary>
        /// <param name="fieldValue">The value to format.</param>
        /// <param name="parameters">Format string parameters.</param>
        /// <returns>The formatted string value.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (parameters.Count > 0)
            {
                var format = parameters[0]?.ToString();
                if (!string.IsNullOrEmpty(format))
                    return string.Format("{0:" + format + "}", fieldValue);

                return fieldValue.ToString();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for extracting substrings from text.
    /// </summary>
    internal class SubEvaluator : IEvaluator
    {
        /// <summary>
        /// Extracts a substring from the field value based on start position and length.
        /// </summary>
        /// <param name="fieldValue">The text to extract from.</param>
        /// <param name="parameters">Start position, length, and optional suffix.</param>
        /// <returns>The extracted substring with optional suffix.</returns>
        public virtual string Evaluate(object fieldValue, List<object> parameters)
        {
            if (parameters.Count >= 1)
            {
                var strValue = fieldValue.ToString();
                if (parameters.Count >= 1)
                {
                    var startObj = parameters[0];
                    if (startObj != null)
                    {
                        try
                        {
                            var startNumber = Convert.ToInt32(startObj);
                            if (startNumber < strValue.Length)
                            {
                                if (parameters.Count >= 2)
                                {
                                    var length = Convert.ToInt32(parameters[1]);
                                    if (length > 0)
                                    {
                                        string subStr;
                                        string posFix = "";
                                        if (startNumber + length < strValue.Length)
                                            subStr = strValue.Substring(startNumber, length);
                                        else
                                            subStr = strValue.Substring(startNumber);

                                        if (parameters.Count >= 3)
                                        {
                                            posFix = parameters[2]?.ToString();
                                        }
                                        return subStr + posFix;
                                    }
                                    else
                                    {
                                        return string.Empty;
                                    }
                                }
                            }
                            else
                            {
                                return string.Empty;
                            }
                        }
                        catch { }
                    }
                }

                return strValue;
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for extracting the left portion of text.
    /// </summary>
    internal class LeftEvaluator : SubEvaluator
    {
        /// <summary>
        /// Extracts the leftmost characters from the field value.
        /// </summary>
        /// <param name="fieldValue">The text to extract from.</param>
        /// <param name="parameters">Length of characters to extract.</param>
        /// <returns>The leftmost characters of the text.</returns>
        public override string Evaluate(object fieldValue, List<object> parameters)
        {
            parameters.Insert(0, 0);
            return base.Evaluate(fieldValue, parameters);
        }
    }

    /// <summary>
    /// Evaluator for extracting the right portion of text.
    /// </summary>
    internal class RightEvaluator : SubEvaluator
    {
        /// <summary>
        /// Extracts the rightmost characters from the field value.
        /// </summary>
        /// <param name="fieldValue">The text to extract from.</param>
        /// <param name="parameters">Length of characters to extract.</param>
        /// <returns>The rightmost characters of the text.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            var str = fieldValue.ToString();
            var strLength = str.Length;
            if (parameters.Count > 0 && int.TryParse(parameters[0].ToString(), out var length))
            {
                if (length > 0)
                    return str.Substring(strLength - length);
                return string.Empty;
            }
            return str;
        }
    }

    /// <summary>
    /// Evaluator for trimming whitespace from text.
    /// </summary>
    internal class TrimEvaluator : IEvaluator
    {
        /// <summary>
        /// Removes leading and trailing whitespace from the field value.
        /// </summary>
        /// <param name="fieldValue">The text to trim.</param>
        /// <param name="parameters">Not used.</param>
        /// <returns>The trimmed text.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                return fieldValue.ToString().Trim();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for converting text to uppercase.
    /// </summary>
    internal class UpperEvaluator : IEvaluator
    {
        /// <summary>
        /// Converts the field value to uppercase.
        /// </summary>
        /// <param name="fieldValue">The text to convert.</param>
        /// <param name="parameters">Not used.</param>
        /// <returns>The uppercase version of the text.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                return fieldValue.ToString().ToUpper();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for converting text to lowercase.
    /// </summary>
    internal class LowerEvaluator : IEvaluator
    {
        /// <summary>
        /// Converts the field value to lowercase.
        /// </summary>
        /// <param name="fieldValue">The text to convert.</param>
        /// <param name="parameters">Not used.</param>
        /// <returns>The lowercase version of the text.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                return fieldValue.ToString().ToLower();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for formatting values as currency.
    /// </summary>
    internal class CurrencyEvaluator : IEvaluator
    {
        /// <summary>
        /// Formats the field value as currency.
        /// </summary>
        /// <param name="fieldValue">The value to format.</param>
        /// <param name="parameters">Not used.</param>
        /// <returns>The formatted currency string.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                return string.Format("{0:C}", fieldValue.ToString());
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for formatting values as percentages.
    /// </summary>
    internal class PercentageEvaluator : IEvaluator
    {
        /// <summary>
        /// Formats the field value as a percentage.
        /// </summary>
        /// <param name="fieldValue">The value to format.</param>
        /// <param name="parameters">Not used.</param>
        /// <returns>The formatted percentage string.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                try
                {
                    var number = Convert.ToDecimal(fieldValue);
                    return string.Format("{0:0.00}", number * 100) + "%";
                }
                catch { }
                return fieldValue.ToString();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for replacing text in strings.
    /// </summary>
    internal class ReplaceEvaluator : IEvaluator
    {
        /// <summary>
        /// Replaces occurrences of text in the field value.
        /// </summary>
        /// <param name="fieldValue">The text to perform replacements in.</param>
        /// <param name="parameters">Search text and replacement text.</param>
        /// <returns>The text with replacements applied.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                var strValue = fieldValue.ToString();
                if (parameters.Count >= 2)
                {
                    var p1 = parameters[0];
                    var p2 = parameters[1];
                    if (p1 != null && p1 != null)
                    {
                        var strP1 = p1.ToString();
                        var strP2 = p2.ToString();
                        try
                        {
                            return Regex.Replace(strValue, strP1, strP2);
                        }
                        catch { }
                    }
                }
                return strValue;
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for conditional expressions.
    /// </summary>
    internal class IfEvaluator : IEvaluator
    {
        /// <summary>
        /// Evaluates a conditional expression and returns the appropriate value.
        /// </summary>
        /// <param name="fieldValue">The value to compare against.</param>
        /// <param name="parameters">Comparison value, true result, and optional false result.</param>
        /// <returns>The result based on the condition.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                if (parameters.Count >= 2)
                {
                    var compareValue = parameters[0];
                    var v1 = parameters[1];

                    if (compareValue != null)
                    {
                        var cp = Utils.CompareObjects(fieldValue, compareValue);
                        if (cp == CompareValue.Eq)
                        {
                            if (v1 != null)
                                return v1.ToString();
                            return string.Empty;
                        }
                        else
                        {
                            if (parameters.Count >= 3)
                            {
                                var v2 = parameters[2];
                                if (v2 != null)
                                    return v2.ToString();
                                return string.Empty;
                            }
                        }
                    }
                }
                return fieldValue.ToString();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for complex conditional expressions.
    /// </summary>
    internal class ConditionEvaluator : IEvaluator
    {
        /// <summary>
        /// Evaluates a complex conditional expression with multiple conditions.
        /// </summary>
        /// <param name="fieldValue">The value to evaluate.</param>
        /// <param name="parameters">Condition parameters including operators and values.</param>
        /// <returns>The result based on the conditions.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                if (parameters.Count >= 3)
                {
                    var v1 = parameters[0];
                    var op = parameters[1]?.ToString();
                    var v2 = parameters[2];

                    if (v1 != null && v2 != null && !string.IsNullOrEmpty(op))
                    {
                        var cp = Utils.CompareObjects(v1, v2);
                        switch (op)
                        {
                            case OperatorName.Gt:
                                if (cp == CompareValue.Gt)
                                {
                                    if (parameters.Count >= 4)
                                        return parameters[3]?.ToString();
                                    return string.Empty;
                                }
                                break;
                            case OperatorName.Lt:
                                if (cp == CompareValue.Lt)
                                {
                                    if (parameters.Count >= 4)
                                        return parameters[3]?.ToString();
                                    return string.Empty;
                                }
                                break;
                            case OperatorName.Eq1:
                            case OperatorName.Eq2:
                                if (cp == CompareValue.Eq)
                                {
                                    if (parameters.Count >= 4)
                                        return parameters[3]?.ToString();
                                    return string.Empty;
                                }
                                break;
                            case OperatorName.Neq1:
                            case OperatorName.Neq2:
                                if (cp != CompareValue.Eq)
                                {
                                    if (parameters.Count >= 4)
                                        return parameters[3]?.ToString();
                                    return string.Empty;
                                }
                                break;
                            case OperatorName.Geq:
                                if (cp == CompareValue.Gt || cp == CompareValue.Eq)
                                {
                                    if (parameters.Count >= 4)
                                        return parameters[3]?.ToString();
                                    return string.Empty;
                                }
                                break;
                            case OperatorName.Leq:
                                if (cp == CompareValue.Lt || cp == CompareValue.Eq)
                                {
                                    if (parameters.Count >= 4)
                                        return parameters[3]?.ToString();
                                    return string.Empty;
                                }
                                break;
                        }
                    }
                }
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for loop operations.
    /// </summary>
    internal class LoopEvaluator : IEvaluator
    {
        /// <summary>
        /// Evaluates a loop operation on a collection.
        /// </summary>
        /// <param name="fieldValue">The collection to iterate over.</param>
        /// <param name="parameters">Loop parameters.</param>
        /// <returns>The result of the loop operation.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            return fieldValue?.ToString() ?? string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for table operations.
    /// </summary>
    internal class TableEvaluator : LoopEvaluator
    {
        /// <summary>
        /// Evaluates a table operation on a collection.
        /// </summary>
        /// <param name="fieldValue">The collection to create a table from.</param>
        /// <param name="parameters">Table parameters.</param>
        /// <returns>The result of the table operation.</returns>
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            return fieldValue?.ToString() ?? string.Empty;
        }
    }

    /// <summary>
    /// Base evaluator for image operations.
    /// </summary>
    internal class ImageEvaluator : IEvaluator
    {
        /// <summary>
        /// Evaluates an image operation.
        /// </summary>
        /// <param name="fieldValue">The image data or URL.</param>
        /// <param name="parameters">Image parameters.</param>
        /// <returns>The result of the image operation.</returns>
        public virtual string Evaluate(object fieldValue, List<object> parameters)
        {
            return fieldValue?.ToString() ?? string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for generating barcode images.
    /// </summary>
    internal class BarCodeEvaluator : ImageEvaluator
    {
        /// <summary>
        /// Generates a barcode image from the field value.
        /// </summary>
        /// <param name="fieldValue">The data to encode in the barcode.</param>
        /// <param name="parameters">Barcode generation parameters.</param>
        /// <returns>The generated barcode image data.</returns>
        public override string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                var barcode = new BarcodeLib.Barcode();
                var height = Constant.DEFAULT_BARCODE_HIGHT;
                var barWidth = Constant.DEFAULT_BARCODE_BARWIDTH;
                var darkColor = Constant.DEFAULT_DARK_COLOR;
                var lightColor = Constant.DEFAULT_LIGHT_COLOR;

                if (parameters.Count >= 1)
                {
                    height = Convert.ToInt32(parameters[0]);
                }
                if (parameters.Count >= 2)
                {
                    barWidth = Convert.ToInt32(parameters[1]);
                }
                if (parameters.Count >= 3)
                {
                    darkColor = SKColor.Parse(parameters[2].ToString());
                }
                if (parameters.Count >= 4)
                {
                    lightColor = SKColor.Parse(parameters[3].ToString());
                }

                var image = barcode.Encode(
                    BarcodeLib.TYPE.CODE128,
                    fieldValue.ToString(),
                    darkColor,
                    lightColor,
                    barWidth,
                    height
                );
                return image.ToString();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for generating QR code images.
    /// </summary>
    internal class QRCodeEvaluator : ImageEvaluator
    {
        /// <summary>
        /// Generates a QR code image from the field value.
        /// </summary>
        /// <param name="fieldValue">The data to encode in the QR code.</param>
        /// <param name="parameters">QR code generation parameters.</param>
        /// <returns>The generated QR code image data.</returns>
        public override string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                var size = Constant.DEFAULT_QRCODE_SIZE;
                var border = Constant.DEFAULT_QRCODE_BORDER;
                var darkColor = Constant.DEFAULT_DARK_COLOR;
                var lightColor = Constant.DEFAULT_LIGHT_COLOR;

                if (parameters.Count >= 1)
                {
                    size = Convert.ToInt32(parameters[0]);
                }
                if (parameters.Count >= 2)
                {
                    border = Convert.ToInt32(parameters[1]);
                }
                if (parameters.Count >= 3)
                {
                    darkColor = SKColor.Parse(parameters[2].ToString());
                }
                if (parameters.Count >= 4)
                {
                    lightColor = SKColor.Parse(parameters[3].ToString());
                }

                var qrCode = new QRCodeGenerator();
                var qrCodeData = qrCode.CreateQrCode(
                    fieldValue.ToString(),
                    QRCodeGenerator.ECCLevel.Q
                );
                var qrCodeImage = new QRCode(qrCodeData);
                var qrCodeBitmap = qrCodeImage.GetGraphic(20, darkColor, lightColor, true);
                return qrCodeBitmap.ToString();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for handling HTML content.
    /// </summary>
    internal class HtmlEvaluator : IEvaluator
    {
        /// <summary>
        /// Processes HTML content for insertion into the document.
        /// </summary>
        /// <param name="fieldValue">The HTML content to process.</param>
        /// <param name="parameters">HTML processing parameters.</param>
        /// <returns>The processed HTML content.</returns>
        public virtual string Evaluate(object fieldValue, List<object> parameters)
        {
            return fieldValue?.ToString() ?? string.Empty;
        }
    }

    /// <summary>
    /// Evaluator for handling Word document content.
    /// </summary>
    internal class WordEvaluator : IEvaluator
    {
        /// <summary>
        /// Processes Word document content for insertion.
        /// </summary>
        /// <param name="fieldValue">The Word document content to process.</param>
        /// <param name="parameters">Word document processing parameters.</param>
        /// <returns>The processed Word document content.</returns>
        public virtual string Evaluate(object fieldValue, List<object> parameters)
        {
            return fieldValue?.ToString() ?? string.Empty;
        }
    }
}
