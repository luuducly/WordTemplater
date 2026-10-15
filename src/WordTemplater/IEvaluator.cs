using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.InkML;
using Newtonsoft.Json.Linq;

namespace WordTemplater
{
  public interface IEvaluator
  {
    /// <summary>
    /// Format the data field value.
    /// </summary>
    /// <param name="fieldValue">
    /// The data field value.
    /// </param>
    /// <param name="parameters">
    /// The other parameters are declared in template file.
    /// </param>
    /// <returns>
    /// The return value will be shown in exported word file.
    /// </returns>
    public string Evaluate(object fieldValue, List<object> parameters);
  }

  internal class DefaultEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null) return string.Empty;

      if (parameters.Count > 0)
      {
        var format = parameters[0]?.ToString();
        if (!string.IsNullOrEmpty(format))
        {
          if (fieldValue is string strVal)
          {
            if (DateTime.TryParse(strVal, out var dt))
              return string.Format("{0:" + format + "}", dt);
            if (decimal.TryParse(strVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num) ||
                decimal.TryParse(strVal, out num))
              return string.Format("{0:" + format + "}", num);
          }
          return string.Format("{0:" + format + "}", fieldValue);
        }

        return fieldValue.ToString() ?? string.Empty;
      }
      return fieldValue.ToString() ?? string.Empty;
    }
  }

  internal class SubEvaluator : IEvaluator
  {
    public virtual string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null) return string.Empty;
      var strValue = fieldValue.ToString() ?? string.Empty;

      if (parameters.Count >= 1)
      {
        var startObj = parameters[0];
        if (startObj != null)
        {
          try
          {
            var startNumber = Convert.ToInt32(startObj);
            if (startNumber < 0) startNumber = 0;
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
                    posFix = parameters[2]?.ToString() ?? "";
                  }
                  return subStr + posFix;
                }
                else
                {
                  return string.Empty;
                }
              }
              return strValue.Substring(startNumber);
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
  }

  internal class LeftEvaluator : SubEvaluator
  {
    public override string Evaluate(object fieldValue, List<object> parameters)
    {
      var clonedParams = new List<object>(parameters);
      clonedParams.Insert(0, 0);
      return base.Evaluate(fieldValue, clonedParams);
    }
  }

  internal class RightEvaluator : SubEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null) return string.Empty;
      var str = fieldValue.ToString() ?? string.Empty;
      var strLength = str.Length;
      if (parameters.Count > 0 && int.TryParse(parameters[0]?.ToString(), out var length))
      {
        if (length <= 0) return string.Empty;
        if (length >= strLength) return str;
        return str.Substring(strLength - length);
      }
      return str;
    }
  }

  internal class TrimEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return fieldValue.ToString()?.Trim() ?? string.Empty;
      }
      return string.Empty;
    }
  }

  internal class UpperEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return fieldValue.ToString()?.ToUpper() ?? string.Empty;
      }
      return string.Empty;
    }
  }

  internal class LowerEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return fieldValue.ToString()?.ToLower() ?? string.Empty;
      }
      return string.Empty;
    }
  }

  internal class CurrencyEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null) return string.Empty;

      decimal number = 0;
      bool parsed = false;
      if (fieldValue is decimal d) { number = d; parsed = true; }
      else if (fieldValue is double db) { number = (decimal)db; parsed = true; }
      else if (fieldValue is float f) { number = (decimal)f; parsed = true; }
      else if (fieldValue is int i) { number = i; parsed = true; }
      else if (fieldValue is long l) { number = l; parsed = true; }
      else
      {
        var str = fieldValue.ToString()?.Trim();
        if (decimal.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number) ||
            decimal.TryParse(str, out number))
        {
          parsed = true;
        }
      }

      if (!parsed) return fieldValue.ToString() ?? string.Empty;

      var culture = System.Globalization.CultureInfo.CurrentCulture;
      string customFormat = "{0:C}";

      if (parameters != null && parameters.Count > 0)
      {
        var param0 = parameters[0]?.ToString()?.Trim();
        if (!string.IsNullOrEmpty(param0))
        {
          try
          {
            culture = System.Globalization.CultureInfo.GetCultureInfo(param0);
          }
          catch
          {
            customFormat = "{0:" + param0 + "}";
          }
        }
      }

      return string.Format(culture, customFormat, number);
    }
  }

  internal class PercentageEvaluator : IEvaluator
  {
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

  internal class ReplaceEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        var strValue = fieldValue.ToString() ?? string.Empty;
        if (parameters.Count >= 2)
        {
          var p1 = parameters[0];
          var p2 = parameters[1];
          if (p1 != null)
          {
            var strP1 = p1.ToString() ?? string.Empty;
            var strP2 = p2?.ToString() ?? string.Empty;
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

  internal class DateEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null) return string.Empty;

      DateTime dt;
      if (fieldValue is DateTime d)
      {
        dt = d;
      }
      else if (fieldValue is DateTimeOffset dto)
      {
        dt = dto.DateTime;
      }
      else
      {
        var str = fieldValue.ToString()?.Trim();
        if (string.IsNullOrEmpty(str)) return string.Empty;
        if (!DateTime.TryParse(str, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt) &&
            !DateTime.TryParse(str, out dt))
        {
          return str;
        }
      }

      var format = parameters.Count > 0 ? parameters[0]?.ToString() : "dd/MM/yyyy";
      if (string.IsNullOrEmpty(format)) format = "dd/MM/yyyy";

      var culture = System.Globalization.CultureInfo.CurrentCulture;
      if (parameters.Count > 1)
      {
        var cultureStr = parameters[1]?.ToString()?.Trim();
        if (!string.IsNullOrEmpty(cultureStr))
        {
          try { culture = System.Globalization.CultureInfo.GetCultureInfo(cultureStr); } catch { }
        }
      }

      return dt.ToString(format, culture);
    }
  }

  internal class NumberEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null) return string.Empty;

      decimal num = 0;
      bool parsed = false;
      if (fieldValue is decimal d) { num = d; parsed = true; }
      else if (fieldValue is double db) { num = (decimal)db; parsed = true; }
      else if (fieldValue is float f) { num = (decimal)f; parsed = true; }
      else if (fieldValue is int i) { num = i; parsed = true; }
      else if (fieldValue is long l) { num = l; parsed = true; }
      else
      {
        var str = fieldValue.ToString()?.Trim();
        if (decimal.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out num) ||
            decimal.TryParse(str, out num))
        {
          parsed = true;
        }
      }

      if (!parsed) return fieldValue.ToString() ?? string.Empty;

      var format = parameters.Count > 0 ? parameters[0]?.ToString() : "#,##0";
      if (string.IsNullOrEmpty(format)) format = "#,##0";

      var culture = System.Globalization.CultureInfo.CurrentCulture;
      if (parameters.Count > 1)
      {
        var cultureStr = parameters[1]?.ToString()?.Trim();
        if (!string.IsNullOrEmpty(cultureStr))
        {
          try { culture = System.Globalization.CultureInfo.GetCultureInfo(cultureStr); } catch { }
        }
      }

      return num.ToString(format, culture);
    }
  }

  internal class DefaultValueEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue == null)
      {
        return parameters.Count > 0 ? (parameters[0]?.ToString() ?? string.Empty) : string.Empty;
      }

      var str = fieldValue.ToString();
      if (string.IsNullOrEmpty(str))
      {
        return parameters.Count > 0 ? (parameters[0]?.ToString() ?? string.Empty) : string.Empty;
      }

      return str;
    }
  }

  internal class PadLeftEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      var str = fieldValue?.ToString() ?? string.Empty;
      if (parameters.Count > 0 && int.TryParse(parameters[0]?.ToString(), out int totalWidth))
      {
        char padChar = ' ';
        if (parameters.Count > 1)
        {
          var p1 = parameters[1]?.ToString();
          if (!string.IsNullOrEmpty(p1)) padChar = p1[0];
        }
        return str.PadLeft(totalWidth, padChar);
      }
      return str;
    }
  }

  internal class PadRightEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      var str = fieldValue?.ToString() ?? string.Empty;
      if (parameters.Count > 0 && int.TryParse(parameters[0]?.ToString(), out int totalWidth))
      {
        char padChar = ' ';
        if (parameters.Count > 1)
        {
          var p1 = parameters[1]?.ToString();
          if (!string.IsNullOrEmpty(p1)) padChar = p1[0];
        }
        return str.PadRight(totalWidth, padChar);
      }
      return str;
    }
  }

  internal class IfEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        if (parameters.Count == 1)
        {
          bool isTrue = IsTruthy(fieldValue);
          return isTrue ? (parameters[0]?.ToString() ?? string.Empty) : string.Empty;
        }
        if (parameters.Count == 2)
        {
          bool isTrue = IsTruthy(fieldValue);
          return isTrue ? (parameters[0]?.ToString() ?? string.Empty) : (parameters[1]?.ToString() ?? string.Empty);
        }
        if (parameters.Count == 4)
        {
          var op = parameters[0]?.ToString();
          var compareValue = parameters[1];
          var v1 = parameters[2]?.ToString() ?? string.Empty;
          var v2 = parameters[3]?.ToString() ?? string.Empty;

          var cp = Utils.CompareObjects(fieldValue, compareValue);
          bool isCorrect = false;
          switch (op)
          {
            case OperatorName.Eq1:
            case OperatorName.Eq2:
              isCorrect = (cp == CompareValue.Eq);
              break;
            case OperatorName.Neq1:
            case OperatorName.Neq2:
              isCorrect = (cp != CompareValue.Eq);
              break;
            case OperatorName.Gt:
              isCorrect = (cp == CompareValue.Gt);
              break;
            case OperatorName.Lt:
              isCorrect = (cp == CompareValue.Lt);
              break;
            case OperatorName.Geq:
              isCorrect = (cp == CompareValue.Gt || cp == CompareValue.Eq);
              break;
            case OperatorName.Leq:
              isCorrect = (cp == CompareValue.Lt || cp == CompareValue.Eq);
              break;
          }
          return isCorrect ? v1 : v2;
        }
        else if (parameters.Count >= 3)
        {
          var compareValue = parameters[0];
          var v1 = parameters[1];
          var v2 = parameters[2];

          if (compareValue != null)
          {
            var cp = Utils.CompareObjects(fieldValue, compareValue);
            if (cp == CompareValue.Eq)
            {
              return v1?.ToString() ?? string.Empty;
            }
            else
            {
              return v2?.ToString() ?? string.Empty;
            }
          }
        }
        return fieldValue.ToString();
      }
      return string.Empty;
    }

    private static bool IsTruthy(object value)
    {
      if (value is bool b) return b;
      if (bool.TryParse(value.ToString(), out bool pb)) return pb;
      if (value is sbyte sb) return sb != 0;
      if (value is byte by) return by != 0;
      if (value is short s) return s != 0;
      if (value is ushort us) return us != 0;
      if (value is int i) return i != 0;
      if (value is uint ui) return ui != 0;
      if (value is long l) return l != 0;
      if (value is ulong ul) return ul != 0;
      if (value is float f) return f != 0.0f;
      if (value is double d) return d != 0.0;
      if (value is decimal dec) return dec != 0m;
      string str = value.ToString()?.Trim() ?? string.Empty;
      return !string.IsNullOrEmpty(str) && !str.Equals("0", StringComparison.OrdinalIgnoreCase) && !str.Equals("false", StringComparison.OrdinalIgnoreCase);
    }
  }

  internal class ConditionEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      bool isCorrect = false;
      if (parameters.Count >= 2)
      {
        var op = parameters[0];
        var compareValue = parameters[1];

        if (op != null)
        {
          var cp = Utils.CompareObjects(ResolveFieldValue(fieldValue), compareValue);
          switch (op)
          {
            case OperatorName.Eq1:
            case OperatorName.Eq2:
              if (cp == CompareValue.Eq) isCorrect = true;
              else isCorrect = false;
              break;
            case OperatorName.Neq1:
            case OperatorName.Neq2:
              if (cp != CompareValue.Eq) isCorrect = true;
              else isCorrect = false;
              break;
            case OperatorName.Gt:
              if (cp == CompareValue.Gt) isCorrect = true;
              else isCorrect = false;
              break;
            case OperatorName.Lt:
              if (cp == CompareValue.Lt) isCorrect = true;
              else isCorrect = false;
              break;
            case OperatorName.Geq:
              if (cp == CompareValue.Gt || cp == CompareValue.Eq) isCorrect = true;
              else isCorrect = false;
              break;
            case OperatorName.Leq:
              if (cp == CompareValue.Lt || cp == CompareValue.Eq) isCorrect = true;
              else isCorrect = false;
              break;
          }

        }
      }
      return isCorrect.ToString();
    }

    private static object ResolveFieldValue(object fieldValue)
    {
      if (fieldValue == null)
        return fieldValue;

      if (fieldValue is JArray jArray)
        return jArray.Count;

      if (fieldValue is Array array)
        return array.Length;

      if (fieldValue is System.Collections.ICollection collection)
        return collection.Count;

      return fieldValue;
    }
  }

  internal class LoopEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      throw new NotImplementedException();
    }
  }

  internal class TableEvaluator : LoopEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      throw new NotImplementedException();
    }
  }

  internal class ImageEvaluator : IEvaluator
  {
    public virtual string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return fieldValue.ToString();
      }
      return string.Empty;
    }
  }

  internal class MermaidEvaluator : ImageEvaluator
  {
    public override string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        try
        {
          var url = "https://kroki.io/mermaid/png";
          using var content = new StringContent(
              fieldValue.ToString() ?? string.Empty,
              Encoding.UTF8,
              "text/plain"
          );

          using var response = Utils.HttpClient.PostAsync(url, content).GetAwaiter().GetResult();
          if (response.IsSuccessStatusCode)
          {
            byte[] imageBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            return Convert.ToBase64String(imageBytes);
          }
        }
        catch
        {
          // Return empty string on network timeout/failure instead of crashing export
        }
      }
      return string.Empty;
    }
  }

  internal class BarCodeEvaluator : ImageEvaluator
  {
    public override string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        var barCode = fieldValue.ToString();
        if (!string.IsNullOrEmpty(barCode))
        {
          var bytes = Utils.GetBarCodeImageBytes(barCode);
          if (bytes != null)
          {
            return Convert.ToBase64String(bytes);
          }

        }
      }
      return string.Empty;
    }
  }

  internal class QRCodeEvaluator : ImageEvaluator
  {
    public override string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        var qCode = fieldValue.ToString();
        if (!string.IsNullOrEmpty(qCode))
        {
          var bytes = Utils.GetQRCodeImageBytes(qCode);
          if (bytes != null)
          {
            return Convert.ToBase64String(bytes);
          }

        }
      }
      return string.Empty;
    }
  }

  internal class HtmlEvaluator : IEvaluator
  {
    public virtual string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return WebUtility.HtmlDecode(fieldValue.ToString());
      }
      return string.Empty;
    }
  }

  internal class WordEvaluator : IEvaluator
  {
    public virtual string Evaluate(object fieldValue, List<object> parameters)
    {
      return fieldValue?.ToString();
    }
  }
}
