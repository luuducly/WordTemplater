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

  internal class SubEvaluator : IEvaluator
  {
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

  internal class LeftEvaluator : SubEvaluator
  {
    public override string Evaluate(object fieldValue, List<object> parameters)
    {
      parameters.Insert(0, 0);
      return base.Evaluate(fieldValue, parameters);
    }
  }

  internal class RightEvaluator : SubEvaluator
  {
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

  internal class TrimEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return fieldValue.ToString().Trim();
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
        return fieldValue.ToString().ToUpper();
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
        return fieldValue.ToString().ToLower();
      }
      return string.Empty;
    }
  }

  internal class CurrencyEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      if (fieldValue != null)
      {
        return string.Format("{0:C}", fieldValue.ToString());
      }
      return string.Empty;
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

  internal class IfEvaluator : IEvaluator
  {
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
              if (v1 != null) return v1.ToString();
              return string.Empty;
            }
            else
            {
              if (parameters.Count >= 3)
              {
                var v2 = parameters[2];
                if (v2 != null) return v2.ToString();
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
              if (cp == (CompareValue.Gt | CompareValue.Eq)) isCorrect = true;
              else isCorrect = false;
              break;
            case OperatorName.Leq:
              if (cp == (CompareValue.Lt | CompareValue.Eq)) isCorrect = true;
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
        using var httpClient = new HttpClient();

        var url = "https://kroki.io/mermaid/png";

        using var content = new StringContent(
            fieldValue.ToString(),
            Encoding.UTF8,
            "text/plain"
        );

        HttpResponseMessage response = httpClient.PostAsync(url, content).Result;

        response.EnsureSuccessStatusCode();

        // Kroki trả về ảnh PNG → đọc dưới dạng byte[]
        byte[] imageBytes = response.Content.ReadAsByteArrayAsync().Result;
        return Convert.ToBase64String(imageBytes);
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
