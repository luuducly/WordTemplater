using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTemplater.Example
{
    internal class UpperCaseFirstLetter : IEvaluator
    {
        public string Evaluate(object fieldValue, List<object> parameters)
        {
            if (fieldValue != null)
            {
                var strValue = fieldValue.ToString() ?? string.Empty;
                TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                return textInfo.ToTitleCase(strValue);
            }
            return string.Empty;
        }
    }
}
