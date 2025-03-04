using Newtonsoft.Json.Linq;
using WordTemplater;
using WordTemplater.Example;

var json = File.ReadAllText("DataSamples\\Data.json");
var data = JObject.Parse(json);
var equationFile = File.ReadAllBytes("Templates\\Equation.docx");
data["Word"] = Convert.ToBase64String(equationFile);

var avatarFile = File.ReadAllBytes("Templates\\Author.jpg");
data["Image"] = Convert.ToBase64String(avatarFile);

var fileName = "Output.docx";
using (var templateStream = File.OpenRead("Templates\\Template.docx"))
{
    using (var wordTemplate = new WordTemplate(templateStream))
    {
        wordTemplate.RegisterEvaluator("customizable", new Number2TextEvaluator());
        wordTemplate.RegisterEvaluator("upperFirstLetter", new UpperCaseFirstLetter());
        using (var exportedStream = wordTemplate.Export(data))
        {
            using (var output = File.Create(fileName))
            {
                exportedStream.CopyTo(output);
            }
        }    
    }
}    

var p = new Process();
p.StartInfo = new ProcessStartInfo(fileName)
{
   UseShellExecute = true
};
p.Start();
