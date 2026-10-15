using System;
using System.Collections.Generic;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class EvaluatorTests
    {
        [Fact]
        public void UpperEvaluator_ShouldConvertToUpperCase()
        {
            var evaluator = new UpperEvaluator();
            var result = evaluator.Evaluate("hello world", new List<object>());
            Assert.Equal("HELLO WORLD", result);
        }

        [Fact]
        public void LowerEvaluator_ShouldConvertToLowerCase()
        {
            var evaluator = new LowerEvaluator();
            var result = evaluator.Evaluate("HELLO WORLD", new List<object>());
            Assert.Equal("hello world", result);
        }

        [Fact]
        public void TrimEvaluator_ShouldTrimWhitespace()
        {
            var evaluator = new TrimEvaluator();
            var result = evaluator.Evaluate("  hello  ", new List<object>());
            Assert.Equal("hello", result);
        }

        [Fact]
        public void SubEvaluator_ShouldSubstringCorrectly()
        {
            var evaluator = new SubEvaluator();
            var result = evaluator.Evaluate("HelloWorld", new List<object> { 0, 5 });
            Assert.Equal("Hello", result);

            var withPostfix = evaluator.Evaluate("HelloWorld", new List<object> { 0, 5, "..." });
            Assert.Equal("Hello...", withPostfix);
        }

        [Fact]
        public void LeftEvaluator_ShouldNotMutateParametersListAcrossCalls()
        {
            var evaluator = new LeftEvaluator();
            var parameters = new List<object> { 4 };
            
            var result1 = evaluator.Evaluate("HelloWorld", parameters);
            Assert.Equal("Hell", result1);

            // Verify second call gets the same length 4, and parameters list was not mutated
            var result2 = evaluator.Evaluate("CSharpTest", parameters);
            Assert.Equal("CSha", result2);
            Assert.Single(parameters);
            Assert.Equal(4, parameters[0]);
        }

        [Fact]
        public void RightEvaluator_ShouldHandleLengthGreaterThanStringLength()
        {
            var evaluator = new RightEvaluator();
            // Should not throw ArgumentOutOfRangeException when length > string.Length
            var result = evaluator.Evaluate("ABC", new List<object> { 5 });
            Assert.Equal("ABC", result);
        }

        [Fact]
        public void CurrencyEvaluator_ShouldFormatNumberOrNumericString()
        {
            var evaluator = new CurrencyEvaluator();
            // Test with explicit culture en-US
            var result1 = evaluator.Evaluate(1234.5m, new List<object> { "en-US" });
            Assert.Equal("$1,234.50", result1);

            // Test with string from JSON in en-US
            var result2 = evaluator.Evaluate("1234.5", new List<object> { "en-US" });
            Assert.Equal("$1,234.50", result2);

            // Test with vi-VN culture
            var result3 = evaluator.Evaluate(1234500, new List<object> { "vi-VN" });
            Assert.Contains("1.234.500", result3.Replace("\u00A0", " "));
        }

        [Fact]
        public void ReplaceEvaluator_ShouldHandleNullReplacementGracefully()
        {
            var evaluator = new ReplaceEvaluator();
            var result = evaluator.Evaluate("abc-def", new List<object> { "-", "" });
            Assert.Equal("abcdef", result);
        }

        [Fact]
        public void ConditionEvaluator_ShouldCompareCorrectly()
        {
            var evaluator = new ConditionEvaluator();

            // Equal
            var eq = evaluator.Evaluate(10, new List<object> { "==", 10 });
            Assert.Equal("True", eq);

            // Greater than
            var gt = evaluator.Evaluate(20, new List<object> { ">", 10 });
            Assert.Equal("True", gt);

            // Less than
            var lt = evaluator.Evaluate(5, new List<object> { "<", 10 });
            Assert.Equal("True", lt);

            // Not equal
            var neq = evaluator.Evaluate("abc", new List<object> { "!=", "xyz" });
            Assert.Equal("True", neq);
        }

        [Fact]
        public void ParameterParser_ShouldHandleQuotesAndCommas()
        {
            var paramStr = "\"arg 1\", 'arg 2', 123, null";
            var parsed = Utils.PaserParametters(paramStr);
            Assert.Equal(4, parsed.Count);
            Assert.Equal("arg 1", parsed[0]);
            Assert.Equal("arg 2", parsed[1]);
            Assert.Equal(123L, parsed[2]);
            Assert.Null(parsed[3]);
        }
    }
}
