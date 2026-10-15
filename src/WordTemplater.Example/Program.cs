using System.Diagnostics;
using Newtonsoft.Json.Linq;
using WordTemplater;
using WordTemplater.Example;

var baseDir = AppContext.BaseDirectory;
string ResolvePath(string relativePath)
{
    if (File.Exists(relativePath)) return relativePath;
    var fromBase = Path.Combine(baseDir, relativePath);
    if (File.Exists(fromBase)) return fromBase;
    return relativePath;
}

var jsonPath = ResolvePath(Path.Combine("DataSamples", "Data.json"));
var json = File.ReadAllText(jsonPath);
var data = JObject.Parse(json);

var equationPath = ResolvePath(Path.Combine("Templates", "Equation.docx"));
var equationFile = File.ReadAllBytes(equationPath);
data["Word"] = Convert.ToBase64String(equationFile);

var avatarPath = ResolvePath(Path.Combine("Templates", "Author.jpg"));
var avatarFile = File.ReadAllBytes(avatarPath);
data["Image"] = Convert.ToBase64String(avatarFile);

var rectPath = ResolvePath(Path.Combine("Templates", "Rectangle.png"));
var rectangleImg = File.ReadAllBytes(rectPath);
data["RectImage"] = Convert.ToBase64String(rectangleImg);

var exportedFileName = "Output.docx";
var templatePath = ResolvePath(Path.Combine("Templates", "Template.docx"));
using (var templateStream = File.OpenRead(templatePath))
{
    using (var wordTemplate = new WordTemplate(templateStream))
    {
        wordTemplate.RegisterEvaluator("customizable", new Number2TextEvaluator());
        wordTemplate.RegisterEvaluator("upperFirstLetter", new UpperCaseFirstLetter());
        using (var exportedStream = wordTemplate.Export(data))
        {
            using (var output = File.Create(exportedFileName))
            {
                exportedStream.CopyTo(output);
            }
        }    
    }
}    

Console.WriteLine($"Exported successfully to {Path.GetFullPath(exportedFileName)}");

if (OperatingSystem.IsWindows())
{
    Process.Start(new ProcessStartInfo(exportedFileName) { UseShellExecute = true });
}
else if (OperatingSystem.IsMacOS())
{
    Process.Start("open", exportedFileName);
}
else if (OperatingSystem.IsLinux())
{
    try
    {
        Process.Start("xdg-open", exportedFileName);
    }
    catch { }
}
