using System.Collections;
using System.Collections.Generic;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DRAW = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;

namespace WordTemplater
{
  public class WordTemplate : IDisposable
  {
    private Stream _sourceStream;
    private readonly byte[] _templateBytes;
    private Dictionary<string, ITemplater> _templaterFactory;
    private ExportOptions _defaultOptions = new ExportOptions();
    private bool _disposed;

    /// <summary>
    /// To init word template to export.
    /// </summary>
    /// <param name="sourceStream">
    /// The template word file stream.
    /// </param>
    /// <exception cref="ArgumentNullException"></exception>
    public WordTemplate(Stream sourceStream)
    {
      if (sourceStream == null)
      {
        throw new ArgumentNullException(nameof(sourceStream));
      }
      _sourceStream = sourceStream;
      if (sourceStream.CanSeek)
        sourceStream.Position = 0;
      using (MemoryStream buffer = new MemoryStream())
      {
        sourceStream.CopyTo(buffer);
        _templateBytes = buffer.ToArray();
      }
      if (sourceStream.CanSeek)
        sourceStream.Position = 0;
      _templaterFactory = new Dictionary<string, ITemplater>();
      this.RegisterTemplater(string.Empty, new DefaultTemplater());
      this.RegisterTemplater(FunctionName.Sub, new SubTemplater());
      this.RegisterTemplater(FunctionName.Left, new LeftTemplater());
      this.RegisterTemplater(FunctionName.Right, new RightTemplater());
      this.RegisterTemplater(FunctionName.Trim, new TrimTemplater());
      this.RegisterTemplater(FunctionName.Upper, new UpperTemplater());
      this.RegisterTemplater(FunctionName.Lower, new LowerTemplater());
      this.RegisterTemplater(FunctionName.If, new IfTemplater());
      this.RegisterTemplater(FunctionName.Currency, new CurrencyTemplater());
      this.RegisterTemplater(FunctionName.Percentage, new PercentageTemplater());
      this.RegisterTemplater(FunctionName.Replace, new ReplaceTemplater());
      this.RegisterTemplater(FunctionName.BarCode, new BarCodeTemplater());
      this.RegisterTemplater(FunctionName.QRCode, new QRCodeTemplater());
      this.RegisterTemplater(FunctionName.Image, new ImageTemplater());
      this.RegisterTemplater(FunctionName.Mermaid, new MermaidTemplater());
      this.RegisterTemplater(FunctionName.Html, new HtmlTemplater());
      this.RegisterTemplater(FunctionName.Word, new WordTemplater());
      this.RegisterTemplater(FunctionName.Date, new DateTemplater());
      this.RegisterTemplater(FunctionName.Number, new NumberTemplater());
      this.RegisterTemplater(FunctionName.DefaultValue, new DefaultValueTemplater());
      this.RegisterTemplater(FunctionName.Coalesce, new DefaultValueTemplater());
      this.RegisterTemplater(FunctionName.PadLeft, new PadLeftTemplater());
      this.RegisterTemplater(FunctionName.PadRight, new PadRightTemplater());
    }

    /// <summary>
    /// To register new format evaluator.
    /// </summary>
    /// <param name="name">
    /// The name of the format evaluator.
    /// </param>
    /// <param name="evaluator">
    /// An instance of IEvaluator.
    /// </param>
    public void RegisterEvaluator(string name, IEvaluator evaluator)
    {
      if (name != null)
      {
        name = name.ToLower();
        _templaterFactory[name] = new CustomTemplater(evaluator);
      }
    }

    private void RegisterTemplater(string name, ITemplater templater)
    {
      if (name != null)
      {
        name = name.ToLower();
        if (!_templaterFactory.ContainsKey(name))
          _templaterFactory.Add(name, templater);
        else
          _templaterFactory[name] = templater;
      }
    }

    /// <summary>
    /// Configures a text watermark for the exported document.
    /// </summary>
    public void SetWatermark(string text, string color = "C0C0C0", double angle = -45, double fontSize = 72, string fontName = "Calibri")
    {
      _defaultOptions.WatermarkText = text;
      _defaultOptions.WatermarkColor = color;
      _defaultOptions.WatermarkAngle = angle;
      _defaultOptions.WatermarkFontSize = fontSize;
      _defaultOptions.WatermarkFontFamily = fontName;
    }

    /// <summary>
    /// Enables automatic zebra striping for tables in the document.
    /// </summary>
    public void EnableTableZebraStriping(string oddColor = "F9FAFB", string? evenColor = null)
    {
      _defaultOptions.TableOddRowColor = oddColor;
      _defaultOptions.TableEvenRowColor = evenColor;
    }

    /// <summary>
    /// To fill data into template file, then export it using custom ExportOptions.
    /// </summary>
    public Stream? Export(object data, ExportOptions options)
    {
      ObjectDisposedException.ThrowIf(_disposed, this);
      if (data != null && options != null)
      {
        Stream targetStream = new MemoryStream();
        targetStream.Write(_templateBytes);
        targetStream.Position = 0;
        using (WordprocessingDocument targetDocument = WordprocessingDocument.Open(targetStream, true))
        {
          if (options.RemoveFallback)
            WordUtils.RemoveFallbackElements(targetDocument);
          if (data is not JObject) data = JObject.FromObject(data);

          if (!string.IsNullOrEmpty(options.WatermarkText))
          {
            WordUtils.AddTextWatermark(
                targetDocument,
                options.WatermarkText,
                options.WatermarkColor,
                options.WatermarkAngle,
                options.WatermarkFontSize,
                options.WatermarkFontFamily);
          }

          List<RenderContext> renderContexts = PrepareRenderContext(targetDocument);
          RenderTemplate(renderContexts, renderContexts.ToList(), data as JObject, options);
          LoopTemplater.FillData(renderContexts, new DataScope(data as JObject));
        }
        targetStream.Position = 0;
        return targetStream;
      }
      return null;
    }

    /// <summary>
    /// To fill data into template file, then export it.
    /// </summary>
    /// <param name="data">
    /// The input data object. Such as JObject or any data model.
    /// </param>
    /// <param name="removeFallBack">
    /// Remove fall back element after exporting. Default value is true.
    /// </param>
    /// <returns>
    /// Return the exported file stream.
    /// </returns>
    public Stream? Export(object data, bool removeFallBack = true)
    {
      var options = new ExportOptions
      {
        RemoveFallback = removeFallBack,
        WatermarkText = _defaultOptions.WatermarkText,
        WatermarkColor = _defaultOptions.WatermarkColor,
        WatermarkAngle = _defaultOptions.WatermarkAngle,
        WatermarkFontSize = _defaultOptions.WatermarkFontSize,
        WatermarkFontFamily = _defaultOptions.WatermarkFontFamily,
        TableOddRowColor = _defaultOptions.TableOddRowColor,
        TableEvenRowColor = _defaultOptions.TableEvenRowColor
      };
      return Export(data, options);
    }


    private List<RenderContext> PrepareRenderContext(WordprocessingDocument document)
    {
      List<RenderContext> renderContexts = new List<RenderContext>();
      MainDocumentPart? mainPart = document.MainDocumentPart;
      if (mainPart == null) return renderContexts;
      var documentPart = mainPart.Document;
      if (documentPart == null) return renderContexts;

      //find bookmark templates in header parts
      foreach (HeaderPart headerPart in mainPart.HeaderParts)
      {
        if (headerPart.Header != null)
          renderContexts.AddRange(PrepareRenderContext(headerPart.Header, headerPart));
      }

      //find bookmark templates in body parts
      if (documentPart.Body != null)
        renderContexts.AddRange(PrepareRenderContext(documentPart.Body, mainPart));

      //find bookmark templates in footer parts
      foreach (FooterPart footerPart in mainPart.FooterParts)
      {
        if (footerPart.Footer != null)
          renderContexts.AddRange(PrepareRenderContext(footerPart.Footer, footerPart));
      }
      return renderContexts;
    }

    private List<RenderContext> PrepareRenderContext(OpenXmlElement element, TypedOpenXmlPart parentPart)
    {
      return PrepareRenderContext(new List<OpenXmlElement>() { element }, parentPart);
    }

    private List<RenderContext> PrepareRenderContext(List<OpenXmlElement> elements, TypedOpenXmlPart parentPart)
    {
      List<RenderContext> rcRootList = new List<RenderContext>();
      Stack<RenderContext> parents = new Stack<RenderContext>();
      Stack<RenderContext> openContexts = new Stack<RenderContext>();
      List<OpenXmlElement> allMergeFieldNodes = new List<OpenXmlElement>();

      foreach (var element in elements)
      {
        allMergeFieldNodes.AddRange(element.Descendants().Where(x => IsMergeFieldNode(x)).ToList());
      }

      for (int i = 0; i < allMergeFieldNodes.Count; i++)
      {
        var mergeFieldNode = allMergeFieldNodes[i];
        var code = GetCode(mergeFieldNode);

        if (code.Contains(FunctionName.EndLoop, StringComparison.OrdinalIgnoreCase)
            || code.Contains(FunctionName.EndTable, StringComparison.OrdinalIgnoreCase)
            || code.Contains(FunctionName.EndIf, StringComparison.OrdinalIgnoreCase)
        )
        {
          if (parents.Count > 0)
            parents.Pop();
          if (openContexts.Count > 0)
            openContexts.Pop().MergeField.EndField = new MergeFieldTemplate(mergeFieldNode);
          continue;
        }

        RenderContext context = new RenderContext();
        context.MergeField = new MergeField();
        context.MergeField.ParentPart = parentPart;
        if (parents.Count > 0)
        {
          context.Parent = parents.Peek();
          context.Parent.ChildNodes.Add(context);
        }
        else
        {
          rcRootList.Add(context);
        }

        context.MergeField.StartField = new MergeFieldTemplate(mergeFieldNode);
        GetFormatTemplate(code, context);
        if (context.Templater != null)
        {
          IEvaluator evaluator = context.Templater.Evaluator;
          if (evaluator != null)
          {
            if (evaluator is LoopEvaluator || evaluator is ConditionEvaluator)
            {
              parents.Push(context);
              openContexts.Push(context);
            }
            else
            {
              continue;
            }
          }
        }
      }
      return rcRootList;
    }

    private bool IsMergeFieldNode(OpenXmlElement x)
    {
      string? fieldCode = null;
      if (x is SimpleField simpleField)
      {
        fieldCode = simpleField.Instruction?.Value;
      }
      else if (x is FieldCode fieldCodeEl)
      {
        fieldCode = fieldCodeEl.Text;
      }

      if (!string.IsNullOrEmpty(fieldCode))
      {
        return fieldCode.Contains("MERGEFORMAT", StringComparison.OrdinalIgnoreCase)
            || fieldCode.Contains("MERGEFIELD", StringComparison.OrdinalIgnoreCase);
      }
      return false;
    }

    private string GetCode(OpenXmlElement node)
    {
      var fieldCode = string.Empty;
      if (node is SimpleField)
      {
        fieldCode = ((SimpleField)node).Instruction.Value;
      }
      else if (node is FieldCode)
      {
        fieldCode = ((FieldCode)node).Text;
      }
      return fieldCode.Trim();
    }


    private void GetFormatTemplate(string code, RenderContext context)
    {
      if (string.IsNullOrEmpty(code)) return;
      var i1 = code.IndexOf(" ");
      var i2 = code.IndexOf(@"\*", StringComparison.OrdinalIgnoreCase);
      if (i2 < 0) i2 = code.IndexOf("MERGEFORMAT", StringComparison.OrdinalIgnoreCase);
      if (i2 < 0) i2 = code.Length;
      if (i1 >= 0 && i2 > i1)
      {
        code = code.Substring(i1, i2 - i1).Trim();
        if (code.StartsWith('"') && code.EndsWith('"') && code.Length >= 2)
        {
          code = code.Substring(1, code.Length - 2);
        }
        var i3 = code.IndexOf('(');

        if (i3 > 0)
        {
          var fmtContent = code.Substring(0, i3);
          var i4 = fmtContent.IndexOf(':');
          var i5 = code.LastIndexOf(')');
          if (i5 < i3) i5 = code.Length - 1;
          var paramContent = code.Substring(i3 + 1, i5 - i3 - 1).Trim();
          if (i4 > 0)
          {
            context.FieldName = fmtContent.Substring(0, i4).Trim();
            var function = fmtContent.Substring(i4 + 1).Trim().ToLower();
            if (_templaterFactory.ContainsKey(function))
              context.Templater = _templaterFactory[function];
            else
              context.Templater = new DefaultTemplater();

            context.Parameters = paramContent;
          }
          else
          {
            var function = fmtContent.Trim().ToLower();
            if (function == FunctionName.If)
            {
              string? op = null;
              if (paramContent.IndexOf(OperatorName.Geq) > 0) op = OperatorName.Geq;
              else if (paramContent.IndexOf(OperatorName.Leq) > 0) op = OperatorName.Leq;
              else if (paramContent.IndexOf(OperatorName.Neq1) > 0) op = OperatorName.Neq1;
              else if (paramContent.IndexOf(OperatorName.Neq2) > 0) op = OperatorName.Neq2;
              else if (paramContent.IndexOf(OperatorName.Gt) > 0) op = OperatorName.Gt;
              else if (paramContent.IndexOf(OperatorName.Lt) > 0) op = OperatorName.Lt;
              else if (paramContent.IndexOf(OperatorName.Eq1) > 0) op = OperatorName.Eq1;
              else if (paramContent.IndexOf(OperatorName.Eq2) > 0) op = OperatorName.Eq2;

              if (!string.IsNullOrEmpty(op))
              {
                int firstComma = paramContent.IndexOf(',');
                if (firstComma > 0)
                {
                  var condPart = paramContent.Substring(0, firstComma).Trim();
                  var branchesPart = paramContent.Substring(firstComma + 1).Trim();
                  var sptContent = condPart.Split(op);
                  context.FieldName = sptContent[0].Trim();
                  var compareVal = sptContent.Length > 1 ? sptContent[1].Trim() : string.Empty;
                  context.Parameters = $"'{op}', {compareVal}, {branchesPart}";
                  context.Templater = _templaterFactory[FunctionName.If];
                }
                else
                {
                  var sptContent = paramContent.Split(op);
                  context.FieldName = sptContent[0].Trim();
                  context.Parameters = sptContent[1].Trim();
                  context.Templater = new ConditionTemplater(op);
                }
              }
              else
              {
                var paramList = Utils.PaserParametters(paramContent);
                if (paramList.Count >= 2)
                {
                  context.FieldName = paramList[0].ToString()!.Trim();
                  int firstComma = paramContent.IndexOf(',');
                  if (firstComma > 0 && firstComma < paramContent.Length - 1)
                  {
                    context.Parameters = paramContent.Substring(firstComma + 1).Trim();
                  }
                  context.Templater = _templaterFactory[FunctionName.If];
                }
                else return;
              }
            }
            else if (function == FunctionName.Loop)
            {
              context.FieldName = paramContent;
              context.Templater = new LoopTemplater();
            }
            else if (function == FunctionName.Table)
            {
              var tableParams = Utils.PaserParametters(paramContent);
              if (tableParams.Count > 1)
              {
                context.FieldName = tableParams[0]?.ToString()?.Trim() ?? string.Empty;
                var tableTemplater = new TableTemplater();
                tableTemplater.OddRowColor = tableParams[1]?.ToString()?.Trim();
                if (tableParams.Count > 2)
                  tableTemplater.EvenRowColor = tableParams[2]?.ToString()?.Trim();
                context.Templater = tableTemplater;
              }
              else
              {
                context.FieldName = paramContent;
                context.Templater = new TableTemplater();
              }
            }
            else if (_templaterFactory.ContainsKey(function))
            {
              var paramList = Utils.PaserParametters(paramContent);
              if (paramList.Count > 0)
              {
                context.FieldName = paramList[0]?.ToString()?.Trim() ?? string.Empty;
                int firstComma = paramContent.IndexOf(',');
                if (firstComma > 0 && firstComma < paramContent.Length - 1)
                {
                  context.Parameters = paramContent.Substring(firstComma + 1).Trim();
                }
                else
                {
                  context.Parameters = string.Empty;
                }
                context.Templater = _templaterFactory[function];
              }
            }
          }
        }
        else
        {
          var i6 = code.IndexOf(':');
          if (i6 > 0)
          {
            context.FieldName = code.Substring(0, i6).Trim();
            var func = code.Substring(i6 + 1).Trim().ToLower();
            if (_templaterFactory.ContainsKey(func))
            {
              context.Templater = _templaterFactory[func];
              context.Parameters = string.Empty;
            }
            else
            {
              context.Templater = new DefaultTemplater();
              context.Parameters = code.Substring(i6 + 1).Trim();
            }
          }
          else
          {
            context.Templater = new DefaultTemplater();
            context.FieldName = code;
          }
        }
      }
    }

    private void RenderTemplate(List<RenderContext> rootContexts, List<RenderContext> renderContexts, JObject dataObj, ExportOptions? options = null)
    {
      if (dataObj != null)
      {
        foreach (var rc in renderContexts)
        {
          var value = Utils.ResolveToken(dataObj, rc.FieldName);
          if (rc.Templater != null && rc.MergeField.StartField != null && rc.MergeField.EndField != null)
          {
            if (rc.Templater.Evaluator is LoopEvaluator && value is JArray)
            {
              RenderTemplate(rootContexts, rc, (JArray)value, options);
            }
            else if (rc.Templater.Evaluator is ConditionEvaluator)
            {
              RenderTemplate(rootContexts, rc.ChildNodes.ToList(), dataObj, options);
            }
          }
        }
      }
    }

    private void RenderTemplate(List<RenderContext> rootContexts, RenderContext rc, JArray arrData, ExportOptions? options = null)
    {
      var template = GetRepeatingTemplate(rc);
      if (arrData.Count == 0)
      {
        foreach (var el in template.TemplateElements)
        {
          el.Remove();
        }
      }
      else
      {
        List<OpenXmlElement> generatedNodes = new List<OpenXmlElement>(template.TemplateElements);

        for (int i = 1; i < arrData.Count; i++)
        {
          generatedNodes.AddRange(template.CloneAndAppendTemplate());
        }

        if (rc.Templater is TableTemplater tableTemplater)
        {
          var oddColor = tableTemplater.OddRowColor ?? options?.TableOddRowColor;
          var evenColor = tableTemplater.EvenRowColor ?? options?.TableEvenRowColor;

          if (!string.IsNullOrEmpty(oddColor) || !string.IsNullOrEmpty(evenColor))
          {
            for (int k = 0; k < generatedNodes.Count; k++)
            {
              if (generatedNodes[k] is WP.TableRow tr)
              {
                if (k % 2 == 1 && !string.IsNullOrEmpty(oddColor))
                {
                  WordUtils.SetRowShading(tr, oddColor);
                }
                else if (k % 2 == 0 && !string.IsNullOrEmpty(evenColor))
                {
                  WordUtils.SetRowShading(tr, evenColor);
                }
              }
            }
          }
        }

        List<RenderContext> allContexts = PrepareRenderContext(generatedNodes, rc.MergeField.ParentPart);

        var firstRC = Find(allContexts, rc);
        if (firstRC != null)
        {
          if (firstRC.Parent != null)
          {
            allContexts = firstRC.Parent.ChildNodes;
          }

          var firstIndex = allContexts.IndexOf(firstRC);
          if (firstIndex >= 0)
          {
            // Item j always lands right after item j-1 (at pos + 1 + j); the recursive call below never modifies this list.
            var pos = rc.Parent != null ? rc.Parent.ChildNodes.IndexOf(rc) : rootContexts.IndexOf(rc);
            for (int j = 0; j < arrData.Count; j++)
            {
              var ct = allContexts[j + firstIndex];
              ct.Index = j;
              if (rc.Parent != null)
              {
                ct.Parent = rc.Parent;
                if (pos > -1)
                  rc.Parent.ChildNodes.Insert(pos + 1 + j, ct);
              }
              else if (rc.Parent == null)
              {
                if (pos > -1)
                  rootContexts.Insert(pos + 1 + j, ct);
              }
              var value = arrData[j];
              if (ct.ChildNodes.Count > 0)
                RenderTemplate(rootContexts, ct.ChildNodes.ToList(), value as JObject, options);
            }

            if (rc.Parent == null)
            {
              rootContexts.Remove(rc);
            }
            else
            {
              rc.Parent.ChildNodes.Remove(rc);
              rc.Parent = null;
            }
          }
        }
      }
    }

    private RepeatingTemplate GetRepeatingTemplate(RenderContext context)
    {
      RepeatingTemplate mfTemplate = new RepeatingTemplate();

      //find the ascendant of both start and end bookmark node
      var mfStart = context.MergeField.StartField.StartNode;
      OpenXmlElement parentNode = mfStart;
      // "x has EndNode as a descendant" == "x is an ancestor of EndNode"; the DOM is not modified while this set is used.
      var endAncestors = new HashSet<OpenXmlElement>(context.MergeField.EndField.EndNode?.Ancestors() ?? Enumerable.Empty<OpenXmlElement>());

      while (parentNode.Parent != null && !endAncestors.Contains(parentNode.Parent))
      {
        parentNode = parentNode.Parent;
      }

      if (parentNode != null)
      {
        OpenXmlElement lastChildNode = null;
        OpenXmlElement startNode = parentNode, endNode = null;

        if (context.Templater.Evaluator is TableEvaluator)
        {
          var tableParentNode = startNode;
          while (tableParentNode != null)
          {
            if (tableParentNode is WP.Table || tableParentNode.Descendants<WP.Table>().Any()
                || tableParentNode == context.MergeField.EndField.EndNode
                || endAncestors.Contains(tableParentNode))
              break;
            tableParentNode = tableParentNode.NextSibling();
          }

          WP.Table tableNode = null;
          if (tableParentNode is WP.Table)
            tableNode = (WP.Table)tableParentNode;
          else
            tableNode = tableParentNode.Descendants<WP.Table>().FirstOrDefault();
          if (tableNode != null)
          {
            TableCell firstCell = null, lastCell = null;
            foreach (var row in tableNode.ChildElements.Where(r => r is WP.TableRow && r.Descendants().Any(x => IsMergeFieldNode(x))))
            {
              if (firstCell == null)
              {
                firstCell = ((WP.TableRow)row).Descendants<TableCell>().FirstOrDefault();
              }
              mfTemplate.TemplateElements.Add(row);
              lastChildNode = row;
            }
            lastCell = ((WP.TableRow)lastChildNode).Descendants<TableCell>().LastOrDefault();

            MoveMergeFieldTo(context.MergeField.StartField, firstCell, true);
            MoveMergeFieldTo(context.MergeField.EndField, lastCell, false);
          }
        }
        else
        {
          var curentNode = startNode;
          while (curentNode != null)
          {
            if (!mfTemplate.TemplateElements.Contains(curentNode))
            {
              mfTemplate.TemplateElements.Add(curentNode);
              lastChildNode = curentNode;
            }

            if (curentNode == context.MergeField.EndField.EndNode || endAncestors.Contains(curentNode))
            {
              endNode = curentNode;
              break;
            }
            curentNode = curentNode.NextSibling();
          }
        }

        if (lastChildNode != null)
        {
          if (lastChildNode.NextSibling() != null)
          {
            mfTemplate.LastNode = lastChildNode.NextSibling();
            mfTemplate.InsertionAnchor = lastChildNode;
          }
          mfTemplate.ParentNode = lastChildNode.Parent;
        }
      }

      return mfTemplate;
    }

    private void MoveMergeFieldTo(MergeFieldTemplate fieldTemplate, TableCell? tableCell, bool forBeginning = true)
    {
      if (fieldTemplate != null && tableCell != null)
      {
        var prg = tableCell.Descendants<WP.Paragraph>().FirstOrDefault();
        if (prg == null)
        {
          prg = new WP.Paragraph();
          tableCell.Append(prg);
        }

        var parentParagraph = fieldTemplate.StartNode.Ancestors<WP.Paragraph>().FirstOrDefault();
        if (forBeginning)
        {
          foreach (var item in fieldTemplate.GetAllElements(true))
          {
            item.Remove();
            prg.InsertAt(item, 0);
          }
        }
        else
        {
          foreach (var item in fieldTemplate.GetAllElements())
          {
            item.Remove();
            prg.Append(item);
          }
        }
        if (parentParagraph != null && parentParagraph.Parent != null && parentParagraph.InnerText.Trim() == string.Empty) parentParagraph.Remove();
      }
    }

    private RenderContext Find(List<RenderContext> renderContexts, RenderContext renderContext)
    {
      foreach (var rc in renderContexts)
      {
        if (rc.MergeField.StartField.StartNode == renderContext.MergeField.StartField.StartNode) return rc;
        if (rc.ChildNodes.Count > 0)
        {
          var rmf = Find(rc.ChildNodes, renderContext);
          if (rmf != null) return rmf;
        }
      }
      return null;
    }

    public void Dispose()
    {
      _disposed = true;
      if (_sourceStream != null)
        _sourceStream.Dispose();
    }
  }

  /// <summary>
  /// Options for customizing document export.
  /// </summary>
  public class ExportOptions
  {
    /// <summary>
    /// Remove fallback elements from template. Default is true.
    /// </summary>
    public bool RemoveFallback { get; set; } = true;

    /// <summary>
    /// Optional watermark text to apply across headers.
    /// </summary>
    public string? WatermarkText { get; set; }

    /// <summary>
    /// Hex color code for watermark text (without #). Default is C0C0C0.
    /// </summary>
    public string WatermarkColor { get; set; } = "C0C0C0";

    /// <summary>
    /// Watermark rotation angle in degrees. Default is -45.
    /// </summary>
    public double WatermarkAngle { get; set; } = -45;

    /// <summary>
    /// Watermark font size in points. Default is 72.
    /// </summary>
    public double WatermarkFontSize { get; set; } = 72;

    /// <summary>
    /// Watermark font family. Default is Calibri.
    /// </summary>
    public string WatermarkFontFamily { get; set; } = "Calibri";

    /// <summary>
    /// Hex color for alternating odd table rows (e.g. F9FAFB).
    /// </summary>
    public string? TableOddRowColor { get; set; }

    /// <summary>
    /// Hex color for alternating even table rows (e.g. FFFFFF).
    /// </summary>
    public string? TableEvenRowColor { get; set; }
  }
}
