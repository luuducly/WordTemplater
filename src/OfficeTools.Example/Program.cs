﻿using Newtonsoft.Json.Linq;
using OfficeTools;
using OfficeTools.Example;

var json1 = File.ReadAllText("DataSamples\\Data1.json");
var data1 = JObject.Parse(json1);
var equationFile = File.ReadAllBytes("Templates\\Equation.docx");
data1["Equation"] = Convert.ToBase64String(equationFile);

var avatarFile = File.ReadAllBytes("Templates\\Author.jpg");
data1["Avatar"] = Convert.ToBase64String(avatarFile);

using (var templateStream = File.OpenRead("Templates\\Template1.docx"))
{
    using (var wordTemplate = new WordTemplate(templateStream))
    {
        wordTemplate.RegisterEvaluator("toText", new Number2TextEvaluator());
        using (var exportedStream = wordTemplate.Export(data1))
        {
            using (var output = File.Create("Output1.docx"))
            {
                exportedStream.CopyTo(output);
            }
        }    
    }
}    