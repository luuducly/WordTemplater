using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class NewFeaturesTests
    {
        [Fact]
        public void DateEvaluator_ShouldFormatDatesCorrectly()
        {
            var evaluator = new DateEvaluator();

            // Date string ISO
            var res1 = evaluator.Evaluate("2026-10-15T14:30:00", new List<object> { "dd/MM/yyyy" });
            Assert.Equal("15/10/2026", res1);

            // DateTime object
            var dt = new DateTime(2026, 12, 25, 9, 0, 0);
            var res2 = evaluator.Evaluate(dt, new List<object> { "yyyy-MM-dd HH:mm" });
            Assert.Equal("2026-12-25 09:00", res2);

            // Culture support
            var res3 = evaluator.Evaluate(new DateTime(2026, 1, 1), new List<object> { "MMMM", "vi-VN" });
            Assert.Contains("tháng", res3.ToLower());
        }

        [Fact]
        public void NumberEvaluator_ShouldFormatNumbersWithCulture()
        {
            var evaluator = new NumberEvaluator();

            // Decimal / double
            var res1 = evaluator.Evaluate(1234567.89, new List<object> { "#,##0.00", "en-US" });
            Assert.Equal("1,234,567.89", res1);

            // String from JSON
            var res2 = evaluator.Evaluate("9876543", new List<object> { "#,##0", "vi-VN" });
            Assert.Equal("9.876.543", res2);
        }

        [Fact]
        public void DefaultValueEvaluator_ShouldFallbackWhenValueIsNullorEmpty()
        {
            var evaluator = new DefaultValueEvaluator();

            var resNull = evaluator.Evaluate(null!, new List<object> { "N/A" });
            Assert.Equal("N/A", resNull);

            var resEmpty = evaluator.Evaluate("", new List<object> { "DefaultVal" });
            Assert.Equal("DefaultVal", resEmpty);

            var resValid = evaluator.Evaluate("ActualValue", new List<object> { "DefaultVal" });
            Assert.Equal("ActualValue", resValid);
        }

        [Fact]
        public void PaddingEvaluators_ShouldPadLeftAndRight()
        {
            var padLeft = new PadLeftEvaluator();
            var resLeft = padLeft.Evaluate(42, new List<object> { 5, "0" });
            Assert.Equal("00042", resLeft);

            var padRight = new PadRightEvaluator();
            var resRight = padRight.Evaluate("INV", new List<object> { 6, "-" });
            Assert.Equal("INV---", resRight);
        }

        [Fact]
        public void ResolveToken_ShouldResolveNestedAndArrayProperties()
        {
            var json = @"{
                ""company"": {
                    ""name"": ""Google LLC"",
                    ""address"": {
                        ""city"": ""Mountain View"",
                        ""zip"": 94043
                    }
                },
                ""employees"": [
                    { ""name"": ""John Doe"", ""role"": ""Engineer"" },
                    { ""name"": ""Jane Smith"", ""role"": ""Architect"" }
                ]
            }";
            var data = JObject.Parse(json);

            // Direct property
            Assert.NotNull(Utils.ResolveToken(data, "company"));

            // Dot notation
            var cityToken = Utils.ResolveToken(data, "company.address.city");
            Assert.NotNull(cityToken);
            Assert.Equal("Mountain View", cityToken.ToString());

            // Case-insensitive dot notation
            var zipToken = Utils.ResolveToken(data, "Company.Address.Zip");
            Assert.NotNull(zipToken);
            Assert.Equal("94043", zipToken.ToString());

            // Array indexing
            var empName = Utils.ResolveToken(data, "employees[0].name");
            Assert.NotNull(empName);
            Assert.Equal("John Doe", empName.ToString());

            // Non-existent path returns null
            Assert.Null(Utils.ResolveToken(data, "company.missing.path"));
        }
    }
}
