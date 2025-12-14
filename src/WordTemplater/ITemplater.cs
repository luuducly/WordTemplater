using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DRAW = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.InkML;
using System.IO;
using System.Xml;

namespace WordTemplater
{
  internal interface ITemplater
  {
    /// <summary>
    /// Indicates the evaluator <see cref="IEvaluator"/>
    /// </summary>
    IEvaluator Evaluator { get; }
    /// <summary>
    /// Replace the template with the corresponding data.
    /// </summary>
    /// <param name="value">The input value</param>
    /// <param name="context">The render context</param>
    void FillData(JToken? value, RenderContext context);
  }

  internal class Templater : ITemplater
  {
    private protected readonly IEvaluator _evaluator;
    IEvaluator ITemplater.Evaluator => _evaluator;
    internal Templater(IEvaluator evaluator)
    {
      _evaluator = evaluator;
    }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      if (value == null)
      {
        context.MergeField.StartField?.RemoveAll();
        return;
      }

      string eval = string.Empty;
      ITemplater templater = context.Templater;

      try
      {
        if (value is JValue jValue)
        {
          var rawValue = jValue.Value;
          if (rawValue != null)
          {
            eval = EvaluateValue(rawValue, templater, context);
          }
        }
        else if (value is JArray jArray)
        {
          var values = jArray
              .Select(v => v is JValue jv && jv.Value != null
                  ? EvaluateValue(jv.Value, templater, context)
                  : string.Empty)
              .Where(s => !string.IsNullOrEmpty(s));

          eval = string.Join(", ", values);
        }
      }
      catch
      {
        eval = value.ToString();
      }

      if (string.IsNullOrEmpty(eval))
      {
        context.MergeField.StartField?.RemoveAll();
        return;
      }

      var textNode = context.MergeField.StartField?.RemoveAllExceptTextNode();
      if (textNode != null)
      {
        textNode.Space = SpaceProcessingModeValues.Preserve;
        textNode.Text = eval;
      }
    }

    private static string EvaluateValue(object value, ITemplater templater, RenderContext context)
    {
      if (templater == null)
        return value.ToString();

      try
      {
        return templater.Evaluator.Evaluate(
            value,
            templater.Evaluator is DefaultEvaluator
                ? new List<object> { context.Parameters }
                : Utils.PaserParametters(context.Parameters)
        );
      }
      catch
      {
        return value.ToString();
      }
    }
  }

  [DebuggerDisplay("[Custom Templater]")]
  internal class CustomTemplater : Templater, ITemplater
  {
    internal CustomTemplater(IEvaluator evaluator) : base(evaluator) { }
  }

  [DebuggerDisplay("[Default Templater]")]
  internal class DefaultTemplater : Templater, ITemplater
  {
    internal DefaultTemplater() : base(new DefaultEvaluator()) { }
  }

  [DebuggerDisplay("[Sub Templater]")]
  internal class SubTemplater : Templater, ITemplater
  {
    internal SubTemplater() : base(new SubEvaluator()) { }
  }

  [DebuggerDisplay("[Left Templater]")]
  internal class LeftTemplater : Templater, ITemplater
  {
    internal LeftTemplater() : base(new LeftEvaluator()) { }
  }

  [DebuggerDisplay("[Right Templater]")]
  internal class RightTemplater : Templater, ITemplater
  {
    internal RightTemplater() : base(new RightEvaluator()) { }
  }

  [DebuggerDisplay("[Trim Templater]")]
  internal class TrimTemplater : Templater, ITemplater
  {
    internal TrimTemplater() : base(new TrimEvaluator()) { }
  }

  [DebuggerDisplay("[Upper Templater]")]
  internal class UpperTemplater : Templater, ITemplater
  {
    internal UpperTemplater() : base(new UpperEvaluator()) { }
  }

  [DebuggerDisplay("[Lower Templater]")]
  internal class LowerTemplater : Templater, ITemplater
  {
    internal LowerTemplater() : base(new LowerEvaluator()) { }
  }

  [DebuggerDisplay("[Currency Templater]")]
  internal class CurrencyTemplater : Templater, ITemplater
  {
    internal CurrencyTemplater() : base(new CurrencyEvaluator()) { }
  }

  [DebuggerDisplay("[Percentage Templater]")]
  internal class PercentageTemplater : Templater, ITemplater
  {
    internal PercentageTemplater() : base(new PercentageEvaluator()) { }
  }

  [DebuggerDisplay("[Replace Templater]")]
  internal class ReplaceTemplater : Templater, ITemplater
  {
    internal ReplaceTemplater() : base(new ReplaceEvaluator()) { }
  }

  [DebuggerDisplay("[If Templater]")]
  internal class IfTemplater : Templater, ITemplater
  {
    internal IfTemplater() : base(new IfEvaluator()) { }
  }

  [DebuggerDisplay("[Condition Templater]")]
  internal class ConditionTemplater : Templater, ITemplater
  {
    internal string _operator { get; set; }
    internal ConditionTemplater(string op) : base(new ConditionEvaluator())
    {
      _operator = op;
    }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      string eval;
      if (value is JValue jvalue)
      {
        var listParam = Utils.PaserParametters(context.Parameters);
        listParam.Insert(0, _operator);
        eval = _evaluator.Evaluate(jvalue, listParam);
      }
      else if (value is JArray array)
      {
        var listParam = Utils.PaserParametters(context.Parameters);
        listParam.Insert(0, _operator);
        eval = _evaluator.Evaluate(array, listParam);
      }
      else
      {
        context.MergeField.StartField?.RemoveAll(true);
        context.MergeField.EndField?.RemoveAll(true);
        return;
      }

      if (string.Compare(true.ToString(), eval, StringComparison.OrdinalIgnoreCase) == 0)
      {
        context.MergeField.StartField?.RemoveAll(true);
        context.MergeField.EndField?.RemoveAll(true);
        LoopTemplater.FillData(context.ChildNodes, context.CurrentItem);
      }
      else
      {
        var start = context.MergeField.StartField.GetAllElements()[0];
        var end = context.MergeField.EndField.GetAllElements(true)[0];
        WordUtils.RemoveFromNodeToNode(start, end);
      }
    }
  }

  [DebuggerDisplay("[Loop Templater]")]
  internal class LoopTemplater : Templater, ITemplater
  {
    internal LoopTemplater() : base(new LoopEvaluator()) { }

    internal LoopTemplater(IEvaluator evaluator) : base(evaluator) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      if (value is JArray)
      {
        var arr = (JArray)value;
        if (arr.Count > 0)
        {
          var arrItem = arr[context.Index];
          if (arrItem is JObject)
          {
            var arrItemJob = arrItem as JObject;
            arrItemJob[Constant.CURRENT_INDEX] = context.Index + 1;
            arrItemJob[Constant.IS_LAST] = (context.Index == arr.Count - 1);
            FillData(context.ChildNodes, arrItemJob);
          }
          else if (arrItem is JValue)
          {
            var jval = new JObject();
            jval[Constant.CURRENT_NODE] = arrItem as JValue;
            jval[Constant.CURRENT_INDEX] = context.Index + 1;
            jval[Constant.IS_LAST] = (context.Index == arr.Count - 1);
            FillData(context.ChildNodes, jval);
          }
          context.MergeField.StartField?.RemoveAll(true);
          context.MergeField.EndField?.RemoveAll(true);
        }
      }
    }

    internal static void FillData(List<RenderContext> renderContexts, JObject data)
    {
      foreach (var context in renderContexts)
      {
        var value = data.GetValue(context.FieldName, StringComparison.OrdinalIgnoreCase);
        context.CurrentItem = data;
        context.Templater?.FillData(value, context);
      }
    }
  }

  [DebuggerDisplay("[Table Templater]")]
  internal class TableTemplater : LoopTemplater, ITemplater
  {
    internal TableTemplater() : base(new TableEvaluator()) { }
  }

  [DebuggerDisplay("[Image Templater]")]
  internal class ImageTemplater : Templater, ITemplater
  {
    internal ImageTemplater() : base(new ImageEvaluator()) { }

    internal ImageTemplater(IEvaluator evaluator) : base(evaluator) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      if (value is JValue)
      {
        var jvalue = ((JValue)value).Value;
        if (jvalue != null)
        {
          var base64Img = _evaluator.Evaluate(jvalue.ToString(), null);
          var stream = new MemoryStream(Convert.FromBase64String(base64Img));

          var drawing = context.MergeField.StartField.StartNode.Ancestors<WP.Drawing>().FirstOrDefault();
          OpenXmlElement imageParentElement = null;
          Size displaySize = null;
          SourceRectangle sourceRectangle = null;
          ShapeTypeValues shapeType = ShapeTypeValues.Rectangle;
          if (drawing != null)
          {
            var run = drawing.Ancestors<WP.Run>().FirstOrDefault();
            DRAW.GraphicData graphicData = null;
            displaySize = GetShapeSize(drawing.Descendants<DRAW.Extents>().FirstOrDefault());
            if (displaySize == null) displaySize = WordUtils.GetImageSize(stream);
            if (displaySize == null)
              goto DONE;
            graphicData = drawing.Descendants<DRAW.GraphicData>().FirstOrDefault();
            if (graphicData != null)
            {
              var geometry = graphicData.Descendants<PresetGeometry>().FirstOrDefault();
              if (geometry != null && geometry.Preset.HasValue)
              {
                shapeType = geometry.Preset.Value;
              }
              graphicData.RemoveAllChildren();
              graphicData.Uri = Constant.PICTURE_NAMESPACE;
              imageParentElement = graphicData;
            }
            else
            {
              if (run == null)
                goto DONE;
              imageParentElement = run;
            }
            Size originalSize = WordUtils.GetImageSize(stream);
            double ratio = Math.Max((double)displaySize.Width / originalSize.Width, (double)displaySize.Height / originalSize.Height);
            Size newImageSize = new Size((long)(originalSize.Width * ratio), (long)(originalSize.Height * ratio));
            int percentVertical = WordUtils.ToThousandPercent((double)(newImageSize.Height - displaySize.Height) / 2 / newImageSize.Height * 100),
              percentHorizontal = WordUtils.ToThousandPercent((double)(newImageSize.Width - displaySize.Width) / 2 / newImageSize.Width * 100);
            sourceRectangle = new SourceRectangle();
            if (percentHorizontal > 0)
              sourceRectangle.Left = sourceRectangle.Right = percentHorizontal;
            if (percentVertical > 0)
              sourceRectangle.Top = sourceRectangle.Bottom = percentVertical;
          }
          else
          {
            Size originalSize = WordUtils.GetImageSize(stream);
            double? percent = GetPercent(context.Parameters);
            if (percent.HasValue)
            {
              (double width, double height) = WordUtils.GetPageSize(context.MergeField.StartField.StartNode);
              double imageWidth = width * percent.Value, imageHeight = height * percent.Value;
              double ratio = Math.Min(imageWidth / originalSize.Width, imageHeight / originalSize.Height);
              displaySize = new Size((long)(originalSize.Width * ratio), (long)(originalSize.Height * ratio));
            }
            else
            {
              displaySize = originalSize;
            }
            (OpenXmlElement drawingElement, OpenXmlElement graphicDataElement) = CreateNewGraphicElement(displaySize);
            imageParentElement = graphicDataElement;
            context.MergeField.StartField.RootNode.InsertBeforeSelf(drawingElement);
          }
          var imgElement = CreateNewPictureElement(stream, context.MergeField.ParentPart, displaySize, shapeType, sourceRectangle);
          imageParentElement.Append(imgElement);
        }
      }
    DONE:
      context.MergeField.StartField.RemoveAll();
    }

    private ImagePart AddImagePart(TypedOpenXmlPart parentPart)
    {
      switch (parentPart)
      {
        case HeaderPart headerPart:
          return headerPart.AddImagePart(ImagePartType.Png);
        case FooterPart footerPart:
          return footerPart.AddImagePart(ImagePartType.Png);
      }
      return ((MainDocumentPart)parentPart).AddImagePart(ImagePartType.Png);
    }

    private (OpenXmlElement drawingElement, OpenXmlElement graphicDataElement) CreateNewGraphicElement(Size size)
    {
      string name = Guid.NewGuid().ToString();
      DRAW.GraphicData graphicDataElement = new DRAW.GraphicData();
      graphicDataElement.Uri = Constant.PICTURE_NAMESPACE;
      var element =
          new WP.Drawing(
              new Inline(
                  new Extent() { Cx = WordUtils.PixelToEmu(size.Width), Cy = WordUtils.PixelToEmu(size.Height) },
                  new EffectExtent()
                  {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                  },
                  new DocProperties()
                  {
                    Id = Utils.GetUintId(),
                    Name = name
                  },
                  new DRAW.NonVisualGraphicFrameDrawingProperties(
                      new DRAW.GraphicFrameLocks() { NoChangeAspect = true }),
                  new DRAW.Graphic(graphicDataElement))
              {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                EditId = Utils.GetRandomHexNumber(8)
              });
      return (element, graphicDataElement);
    }

    private PIC.Picture CreateNewPictureElement(MemoryStream stream, TypedOpenXmlPart parentPart, Size size, ShapeTypeValues type, SourceRectangle sourceRect = null)
    {
      var imagePart = AddImagePart(parentPart);
      var imageId = parentPart.GetIdOfPart(imagePart);
      imagePart.FeedData(stream);
      PIC.BlipFill blipFill = new PIC.BlipFill(
          new DRAW.Blip(
            new DRAW.BlipExtensionList())
          {
            CompressionState = DRAW.BlipCompressionValues.Print,
            Embed = imageId
          })
      {
        RotateWithShape = true
      };
      if (sourceRect != null)
      {
        blipFill.Append(sourceRect);
      }
      blipFill.Append(new DRAW.Stretch());
      PIC.Picture element = new PIC.Picture(
        new PIC.NonVisualPictureProperties(
          new PIC.NonVisualDrawingProperties()
          {
            Id = Utils.GetUintId(),
            Name = string.Format(Constant.DEFAULT_IMAGE_FILE_NAME, Guid.NewGuid().ToString())
          },
          new PIC.NonVisualPictureDrawingProperties()),
        blipFill,
        new PIC.ShapeProperties(
          new DRAW.Transform2D(
            new DRAW.Offset() { X = 0L, Y = 0L },
            new DRAW.Extents() { Cx = WordUtils.PixelToEmu(size.Width), Cy = WordUtils.PixelToEmu(size.Height) }),
          new DRAW.PresetGeometry(
            new DRAW.AdjustValueList())
          {
            Preset = type
          }));
      return element;
    }

    private Size GetShapeSize(DRAW.Extents extents)
    {
      if (extents != null)
      {
        Int64Value? w = extents.Cx;
        Int64Value? h = extents.Cy;
        if (w.HasValue && h.HasValue)
          return new Size((long)WordUtils.EmuToPixels(w.Value), (long)WordUtils.EmuToPixels(h.Value));
      }
      return null;
    }

    private double? GetPercent(string parameters)
    {
      if (parameters == null)
        return null;
      parameters = parameters.Trim();
      if (parameters.Length == 0)
        return null;
      return Utils.GetDouble(parameters);
    }

    private DRAW.ShapeTypeValues GetEnumShapeType(string shapeType)
    {
      switch (shapeType)
      {
        case "line": return ShapeTypeValues.Line;
        case "lineInv": return ShapeTypeValues.LineInverse;
        case "triangle": return ShapeTypeValues.Triangle;
        case "rtTriangle": return ShapeTypeValues.RightTriangle;
        case "rect": return ShapeTypeValues.Rectangle;
        case "diamond": return ShapeTypeValues.Diamond;
        case "parallelogram": return ShapeTypeValues.Parallelogram;
        case "trapezoid": return ShapeTypeValues.Trapezoid;
        case "nonIsoscelesTrapezoid": return ShapeTypeValues.NonIsoscelesTrapezoid;
        case "pentagon": return ShapeTypeValues.Pentagon;
        case "hexagon": return ShapeTypeValues.Hexagon;
        case "heptagon": return ShapeTypeValues.Heptagon;
        case "octagon": return ShapeTypeValues.Octagon;
        case "decagon": return ShapeTypeValues.Decagon;
        case "dodecagon": return ShapeTypeValues.Dodecagon;
        case "star4": return ShapeTypeValues.Star4;
        case "star5": return ShapeTypeValues.Star5;
        case "star6": return ShapeTypeValues.Star6;
        case "star7": return ShapeTypeValues.Star7;
        case "star8": return ShapeTypeValues.Star8;
        case "star10": return ShapeTypeValues.Star10;
        case "star12": return ShapeTypeValues.Star12;
        case "star16": return ShapeTypeValues.Star16;
        case "star24": return ShapeTypeValues.Star24;
        case "star32": return ShapeTypeValues.Star32;
        case "roundRect": return ShapeTypeValues.RoundRectangle;
        case "round1Rect": return ShapeTypeValues.Round1Rectangle;
        case "round2SameRect": return ShapeTypeValues.Round2SameRectangle;
        case "round2DiagRect": return ShapeTypeValues.Round2DiagonalRectangle;
        case "snipRoundRect": return ShapeTypeValues.SnipRoundRectangle;
        case "snip1Rect": return ShapeTypeValues.Snip1Rectangle;
        case "snip2SameRect": return ShapeTypeValues.Snip2SameRectangle;
        case "snip2DiagRect": return ShapeTypeValues.Snip2DiagonalRectangle;
        case "plaque": return ShapeTypeValues.Plaque;
        case "ellipse": return ShapeTypeValues.Ellipse;
        case "teardrop": return ShapeTypeValues.Teardrop;
        case "homePlate": return ShapeTypeValues.HomePlate;
        case "chevron": return ShapeTypeValues.Chevron;
        case "pieWedge": return ShapeTypeValues.PieWedge;
        case "pie": return ShapeTypeValues.Pie;
        default: return ShapeTypeValues.Rectangle;
      }
    }
  }

  internal class MermaidTemplater : ImageTemplater, ITemplater
  {
    internal MermaidTemplater() : base(new MermaidEvaluator()) { }
  }

  [DebuggerDisplay("[BarCode Templater]")]
  internal class BarCodeTemplater : ImageTemplater, ITemplater
  {
    internal BarCodeTemplater() : base(new BarCodeEvaluator()) { }
  }

  internal class QRCodeTemplater : ImageTemplater, ITemplater
  {
    internal QRCodeTemplater() : base(new QRCodeEvaluator()) { }
  }

  [DebuggerDisplay("[Html Templater]")]
  internal class HtmlTemplater : Templater, ITemplater
  {
    internal HtmlTemplater() : base(new HtmlEvaluator()) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      bool isRemovedMergeField = false;
      if (value is JValue)
      {
        var jvalue = ((JValue)value).Value;
        var eval = _evaluator.Evaluate(jvalue, null);
        MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(string.Format(Constant.HTML_PATTERN, eval)));
        AlternativeFormatImportPart formatImportPart = null;
        if (context.MergeField.ParentPart is MainDocumentPart)
          formatImportPart = ((MainDocumentPart)context.MergeField.ParentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
        else if (context.MergeField.ParentPart is HeaderPart)
          formatImportPart = ((HeaderPart)context.MergeField.ParentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
        else if (context.MergeField.ParentPart is FooterPart)
          formatImportPart = ((FooterPart)context.MergeField.ParentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);

        if (formatImportPart != null)
        {
          formatImportPart.FeedData(stream);
          AltChunk altChunk = new AltChunk();
          altChunk.Id = context.MergeField.ParentPart.GetIdOfPart(formatImportPart);
          var node = context.MergeField.StartField?.RemoveAllExceptTextNode();
          if (node != null && node.Parent != null && node.Parent.Parent != null)
          {
            node.Parent.InsertAfterSelf(new WP.Run(altChunk));
            node.Parent.Remove();
            isRemovedMergeField = true;
          }
        }
        stream.Dispose();
      }
      if (!isRemovedMergeField)
        context.MergeField.StartField?.RemoveAll();
    }
  }

  [DebuggerDisplay("[Word Templater]")]
  internal class WordTemplater : Templater, ITemplater
  {
    internal WordTemplater() : base(new WordEvaluator()) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      if (value is JValue)
      {
        if (context.MergeField.ParentPart is MainDocumentPart)
        {
          var body = ((MainDocumentPart)context.MergeField.ParentPart).Document.Body;
          var startNodeToInsert = context.MergeField.StartField.StartNode.Ancestors().FirstOrDefault(a => a.Parent == body);
          if (startNodeToInsert != null)
          {
            var jvalue = ((JValue)value).Value;
            var eval = _evaluator.Evaluate(jvalue, null);
            Stream stream = new MemoryStream(Convert.FromBase64String(eval));
            var wordDocument = WordprocessingDocument.Open(stream, false);
            Dictionary<string, string> mappingRID = new Dictionary<string, string>();
            foreach (var p in wordDocument.MainDocumentPart.Parts)
            {
              //ignore header and footer data
              if (p.OpenXmlPart is HeaderPart or FooterPart)
              {
                continue;
              }

              try
              {
                var rId = Utils.GetUniqueStringID();
                context.MergeField.ParentPart.AddPart(p.OpenXmlPart, rId);
                mappingRID.Add(p.RelationshipId, rId);
              }
              catch { }
            }

            foreach (var el in wordDocument.MainDocumentPart.Document.Body.Elements())
            {
              if (el is SectionProperties) continue;
              var newEl = el.CloneNode(true);
              var subEls = newEl.Descendants().Where(x => { return x.GetAttributes().Where(a => mappingRID.ContainsKey(a.Value)).FirstOrDefault().LocalName != null; });

              foreach (var x in subEls)
              {
                var att = x.GetAttributes().Where(a => mappingRID.ContainsKey(a.Value)).FirstOrDefault();
                var newAttr = new OpenXmlAttribute(att.Prefix, att.LocalName, att.NamespaceUri, mappingRID[att.Value]);
                x.SetAttribute(newAttr);
              }

              if (startNodeToInsert is WP.Paragraph && newEl is WP.Paragraph)
              {
                var paraProp = startNodeToInsert.Elements<WP.ParagraphProperties>().FirstOrDefault();
                var oldParaProp = newEl.Elements<WP.ParagraphProperties>().FirstOrDefault();
                if (paraProp != null)
                {
                  var newParaProp = new WP.ParagraphProperties();
                  if (oldParaProp != null)
                  {
                    foreach (var item in oldParaProp.ChildElements)
                    {
                      if (!(item is Indentation))
                        newParaProp.Append(item.CloneNode(true));
                    }

                    oldParaProp.Remove();
                  }

                  foreach (var item in paraProp.ChildElements)
                  {
                    if (!newParaProp.Elements().Any(x => x.GetType() == item.GetType()))
                      newParaProp.Append(item.CloneNode(true));
                  }

                  newEl.InsertAt(newParaProp, 0);
                }
              }

              startNodeToInsert.InsertBeforeSelf(newEl);
            }

            wordDocument.Dispose();
          }
        }
      }
      context.MergeField.StartField?.RemoveAll(true);
    }
  }
}
