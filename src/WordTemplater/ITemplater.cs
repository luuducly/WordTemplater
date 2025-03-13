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
      if (value is JValue)
      {
        var jvalue = ((JValue)value).Value;
        if (jvalue != null)
        {
          string eval = "";
          ITemplater templater = context.Templater;
          if (templater != null)
          {
            try
            {
              eval = templater.Evaluator.Evaluate(jvalue, templater.Evaluator is DefaultEvaluator ? new List<object>() { context.Parameters } : Utils.PaserParametters(context.Parameters));
            }
            catch
            {
              eval = jvalue.ToString();
            }
          }
          else
          {
            eval = jvalue.ToString();
          }

          var textNode = context.MergeField.StartField?.RemoveAllExceptTextNode();
          if (textNode != null)
          {
            textNode.Space = SpaceProcessingModeValues.Preserve;
            textNode.Text = eval;
          }
          else
          {

          }
        }
        else
        {
          context.MergeField.StartField?.RemoveAll();
        }
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
    internal ConditionTemplater() : base(new ConditionEvaluator()) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
      if (value is JValue)
      {
        var jvalue = ((JValue)value).Value;
        var listParam = Utils.PaserParametters(context.Parameters);
        listParam.Insert(0, context.Operator);
        var eval = _evaluator.Evaluate(jvalue, listParam);
        if (string.Compare(true.ToString(), eval, StringComparison.OrdinalIgnoreCase) == 0)
        {
          context.MergeField.StartField?.RemoveAll(true);
          context.MergeField.EndField?.RemoveAll(true);
        }
        else
        {
          var start = context.MergeField.StartField.GetAllElements()[0];
          var end = context.MergeField.EndField.GetAllElements(true)[0];
          WordUtils.RemoveFromNodeToNode(start, end);
        }
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

    internal ImageTemplater(IEvaluator evaluator) : base (evaluator) { }

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
          if (drawing != null)
          {
            var run = drawing.Ancestors<WP.Run>().FirstOrDefault();

            DRAW.GraphicData graphicData = null;
            Size frame = GetShapeSize(drawing.Descendants<DRAW.Extents>().FirstOrDefault());
            if (frame == null) frame = Utils.GetImageSize(stream);
            if (frame != null)
            {
              WP.Run? pRun = drawing.Ancestors<WP.Run>().FirstOrDefault();
              if (pRun != null)
              {
                graphicData = drawing?.Descendants<DRAW.GraphicData>().FirstOrDefault();
                if (graphicData != null)
                {
                  graphicData.RemoveAllChildren();
                  graphicData.Uri = Constant.PICTURE_NAMESPACE;
                }
              }
            }

            var imgElement = CreateNewPictureElement(Guid.NewGuid().ToString(), 0, 0);
            var imagePart = AddImagePart(context.MergeField.ParentPart);
            var imageId = context.MergeField.ParentPart.GetIdOfPart(imagePart);
            imagePart.FeedData(stream);
            UpdateImageIdAndSize(imgElement, imageId, frame);
            if (graphicData != null)
            {
              graphicData.Append(imgElement);
            }
            else if (run != null)
            {
              run.Append(imgElement);
            }
          }
          else
          {
            OpenXmlElement imgElement = CreateNewPictureElement(Guid.NewGuid().ToString(), 0, 0);
            var frame = Utils.GetImageSize(stream);
            if (frame != null)
            {
              var imagePart = AddImagePart(context.MergeField.ParentPart);
              var imageId = context.MergeField.ParentPart.GetIdOfPart(imagePart);
              imagePart.FeedData(stream);
              UpdateImageIdAndSize(imgElement, imageId, frame);
              imgElement = CreateNewDrawingElement(imgElement, frame);
            }
            var start = context.MergeField.StartField.GetAllElements()[0];
            start.InsertBeforeSelf(imgElement);
          }
        }
      }
      context.MergeField.StartField.RemoveAll();
    }

    private ImagePart AddImagePart(TypedOpenXmlPart parentPart)
    {
      ImagePart imagePart = null;
      if (parentPart is HeaderPart)
      {
        imagePart = ((HeaderPart)parentPart).AddImagePart(ImagePartType.Png);
      }
      else if (parentPart is MainDocumentPart)
      {
        imagePart = ((MainDocumentPart)parentPart).AddImagePart(ImagePartType.Png);
      }
      else if (parentPart is FooterPart)
      {
        imagePart = ((FooterPart)parentPart).AddImagePart(ImagePartType.Png);
      }
      return imagePart;
    }

    private OpenXmlElement CreateNewDrawingElement(OpenXmlElement image, Size size)
    {
      string name = Guid.NewGuid().ToString();
      var element =
          new WP.Drawing(
              new Inline(
                  new Extent() { Cx = size.Width, Cy = size.Height },
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
                  new DRAW.Graphic(
                      new DRAW.GraphicData(
                              image
                          )
                      { Uri = Constant.PICTURE_NAMESPACE })
              )
              {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                EditId = Utils.GetRandomHexNumber(8)
              });
      return element;
    }

    private PIC.Picture CreateNewPictureElement(string fileName, long width, long height)
    {
      return new PIC.Picture(
          new PIC.NonVisualPictureProperties(
              new PIC.NonVisualDrawingProperties()
              {
                Id = Utils.GetUintId(),
                Name = string.Format(Constant.DEFAULT_IMAGE_FILE_NAME, fileName)
              },
              new PIC.NonVisualPictureDrawingProperties()),
          new PIC.BlipFill(
              new DRAW.Blip(
                  new DRAW.BlipExtensionList()
              )
              {
                CompressionState =
                      DRAW.BlipCompressionValues.Print
              },
              new DRAW.Stretch(
                  new DRAW.FillRectangle())),
          new PIC.ShapeProperties(
              new DRAW.Transform2D(
                  new DRAW.Offset() { X = 0L, Y = 0L },
                  new DRAW.Extents() { Cx = width, Cy = height }),
              new DRAW.PresetGeometry(
                      new DRAW.AdjustValueList()
                  )
              { Preset = DRAW.ShapeTypeValues.Rectangle }));
    }

    private void UpdateImageIdAndSize(OpenXmlElement element, string imageId, Size size)
    {
      if (element != null)
      {
        var blip = element.Descendants<DRAW.Blip>().FirstOrDefault();
        if (blip != null) blip.Embed = imageId;
        var extents = element.Descendants<DRAW.Extents>().FirstOrDefault();
        if (extents != null)
        {
          extents.Cx = size.Width;
          extents.Cy = size.Height;
        }
      }
    }

    private Size GetShapeSize(DRAW.Extents extents)
    {
      if (extents != null)
      {
        Int64Value? w = extents.Cx;
        Int64Value? h = extents.Cy;
        if (w.HasValue && h.HasValue)
          return new Size(w.Value, h.Value);
      }
      return null;
    }
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
