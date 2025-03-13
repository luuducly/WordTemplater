using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTemplater.Example
{
  internal class Number2TextEvaluator : IEvaluator
  {
    public string Evaluate(object fieldValue, List<object> parameters)
    {
      double.TryParse(fieldValue.ToString(), out var number);
      return Number2Text.So_chu(number);
    }
  }
}
