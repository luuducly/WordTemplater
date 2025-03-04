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
        private List<RenderContext> _renderContexts;
        private Stream _sourceStream;
        private Dictionary<string, IEvaluator> _evaluatorFactory;

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
            _renderContexts = new List<RenderContext>();
            _evaluatorFactory = new Dictionary<string, IEvaluator>();
            this.RegisterEvaluator(string.Empty, new DefaultEvaluator());
            this.RegisterEvaluator(FunctionName.Sub, new SubEvaluator());
            this.RegisterEvaluator(FunctionName.Left, new LeftEvaluator());
            this.RegisterEvaluator(FunctionName.Right, new RightEvaluator());
            this.RegisterEvaluator(FunctionName.Trim, new TrimEvaluator());
            this.RegisterEvaluator(FunctionName.Upper, new UpperEvaluator());
            this.RegisterEvaluator(FunctionName.Lower, new LowerEvaluator());
            this.RegisterEvaluator(FunctionName.If, new IfEvaluator());
            this.RegisterEvaluator(FunctionName.Currency, new CurrencyEvaluator());
            this.RegisterEvaluator(FunctionName.Percentage, new PercentageEvaluator());
            this.RegisterEvaluator(FunctionName.Replace, new ReplaceEvaluator());
            this.RegisterEvaluator(FunctionName.BarCode, new BarCodeEvaluator());
            this.RegisterEvaluator(FunctionName.QRCode, new QRCodeEvaluator());
            this.RegisterEvaluator(FunctionName.Image, new ImageEvaluator());
            this.RegisterEvaluator(FunctionName.Html, new HtmlEvaluator());
            this.RegisterEvaluator(FunctionName.Word, new WordEvaluator());
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
                if (!_evaluatorFactory.ContainsKey(name))
                    _evaluatorFactory.Add(name, evaluator);
                else
                    _evaluatorFactory[name] = evaluator;
            }
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
        public Stream Export(object data, bool removeFallBack = true)
        {
            if (data != null)
            {
                Stream targetStream = Utils.CloneStream(_sourceStream);
                if (targetStream != null)
                {
                    using (WordprocessingDocument targetDocument = WordprocessingDocument.Open(targetStream, true))
                    {
                        if (removeFallBack)
                            RemoveFallbackElements(targetDocument);
                        if (data is not JObject) data = JObject.FromObject(data);

                        PrepareRenderContext(targetDocument);
                        RenderTemplate(_renderContexts.ToList(), data as JObject);
                        FillData(_renderContexts, data as JObject);
                    }
                }
                targetStream.Position = 0;
                return targetStream;
            }
            return null;
        }


        private void PrepareRenderContext(WordprocessingDocument document)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            var documentPart = mainPart.Document;

            //find bookmark templates in header parts
            foreach (HeaderPart headerPart in mainPart.HeaderParts)
            {
                _renderContexts.AddRange(PrepareRenderContext(headerPart.Header, headerPart));
            }

            //find bookmark templates in body parts
            _renderContexts.AddRange(PrepareRenderContext(documentPart.Body, mainPart));

            //find bookmark templates in footer parts
            foreach (FooterPart footerPart in mainPart.FooterParts)
            {
                _renderContexts.AddRange(PrepareRenderContext(footerPart.Footer, footerPart));
            }
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

                if (code.Contains(FunctionName.EndIf, StringComparison.OrdinalIgnoreCase))
                {
                    if (openContexts.Count > 0)
                        openContexts.Pop().MergeField.EndField = new MergeFieldTemplate(mergeFieldNode);
                    continue;
                }

                if (code.Contains(FunctionName.EndLoop, StringComparison.OrdinalIgnoreCase) || code.Contains(FunctionName.EndTable, StringComparison.OrdinalIgnoreCase))
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
                if (context.Evaluator != null && (context.Evaluator is LoopEvaluator || context.Evaluator is ConditionEvaluator))
                {
                    if (context.Evaluator is LoopEvaluator)
                        parents.Push(context);

                    openContexts.Push(context);
                }
            }
            return rcRootList;
        }

        private bool IsMergeFieldNode(OpenXmlElement x)
        {
            if (x is SimpleField)
            {
                var simpleField = (SimpleField)x;
                var fieldCode = simpleField.Instruction.Value;
                if (!string.IsNullOrEmpty(fieldCode) && fieldCode.Contains(Constant.MERGEFORMAT)) return true;
            }
            else if (x is FieldCode)
            {
                var fieldCode = ((FieldCode)x).Text;
                if (!string.IsNullOrEmpty(fieldCode) && fieldCode.Contains(Constant.MERGEFORMAT)) return true;
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
            var i2 = code.IndexOf(Constant.MERGEFORMAT);
            if (i1 >= 0 && i2 >= 0 && i2 > i1)
            {
                code = code.Substring(i1, i2 - i1).Trim();
                if (code.StartsWith('"') && code.EndsWith('"'))
                {
                    code = code.Substring(1);
                    code = code.Substring(0, code.Length - 1);
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
                        if (_evaluatorFactory.ContainsKey(function))
                            context.Evaluator = _evaluatorFactory[function];
                        else
                            context.Evaluator = _evaluatorFactory[FunctionName.Default];

                        context.Parameters = paramContent;
                    }
                    else
                    {
                        var function = fmtContent.Trim().ToLower();
                        if (function == FunctionName.If)
                        {
                            context.Evaluator = new ConditionEvaluator();
                            if (paramContent.IndexOf(OperatorName.Geq) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Geq);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Geq;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Leq) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Leq);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Leq;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Neq1) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Neq1);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Neq1;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Neq2) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Neq2);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Neq2;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Gt) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Gt);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Gt;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Lt) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Lt);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Lt;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Eq1) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Eq1);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Eq1;
                                context.Parameters = sptContent[1].Trim();
                            }
                            else if (paramContent.IndexOf(OperatorName.Eq2) > 0)
                            {
                                var sptContent = paramContent.Split(OperatorName.Eq2);
                                context.FieldName = sptContent[0].Trim();
                                context.Operator = OperatorName.Eq2;
                                context.Parameters = sptContent[1].Trim();
                            }
                        }
                        else if (function == FunctionName.Loop)
                        {
                            context.FieldName = paramContent;
                            context.Evaluator = new LoopEvaluator();
                        }
                        else if (function == FunctionName.Table)
                        {
                            context.FieldName = paramContent;
                            context.Evaluator = new TableEvaluator();
                        }
                    }
                }
                else
                {
                    var i6 = code.IndexOf(':');
                    if (i6 > 0)
                    {
                        context.FieldName = code.Substring(0, i6).Trim();
                        context.Parameters = code.Substring(i6 + 1).Trim();
                        context.Evaluator = _evaluatorFactory[FunctionName.Default];
                    }
                    else
                    {
                        context.FieldName = code;
                    }
                }
            }
        }

        private void RenderTemplate(List<RenderContext> renderContexts, JObject dataObj)
        {
            if (dataObj != null)
            {
                foreach (var rc in renderContexts)
                {
                    var value = dataObj.GetValue(rc.FieldName, StringComparison.OrdinalIgnoreCase);
                    if (rc.Evaluator is LoopEvaluator && value is JArray && rc.MergeField.StartField != null && rc.MergeField.EndField != null)
                    {
                        RenderTemplate(rc, (JArray)value);
                    }
                }
            }
        }

        private void RenderTemplate(RenderContext rc, JArray arrData)
        {
            var template = GetRepeatingTemplate(rc);
            if (arrData.Count > 1)
            {
                List<OpenXmlElement> generatedNodes = new List<OpenXmlElement>();
                generatedNodes.AddRange(template.TemplateElements);

                for (int i = 1; i < arrData.Count; i++)
                {
                    generatedNodes.AddRange(template.CloneAndAppendTemplate());
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
                        var lastRC = rc;
                        for (int j = 0; j < arrData.Count; j++)
                        {
                            var ct = allContexts[j + firstIndex];
                            ct.Index = j;
                            if (rc.Parent != null)
                            {
                                ct.Parent = rc.Parent;
                                var pos = rc.Parent.ChildNodes.IndexOf(lastRC);
                                if (pos > -1)
                                {
                                    if (pos < rc.Parent.ChildNodes.Count - 1)
                                        rc.Parent.ChildNodes.Insert(pos + 1, ct);
                                    else
                                        rc.Parent.ChildNodes.Add(ct);
                                    lastRC = ct;
                                }
                            }
                            else if (rc.Parent == null)
                            {
                                var pos = _renderContexts.IndexOf(lastRC);
                                if (pos > -1)
                                {
                                    if (pos < _renderContexts.Count - 1)
                                        _renderContexts.Insert(pos + 1, ct);
                                    else
                                        _renderContexts.Add(ct);
                                    lastRC = ct;
                                }
                            }
                            var value = arrData[j];
                            if (ct.ChildNodes.Count > 0)
                                RenderTemplate(ct.ChildNodes.ToList(), value as JObject);
                        }

                        if (rc.Parent == null)
                        {
                            _renderContexts.Remove(rc);
                        }
                        else
                        {
                            rc.Parent.ChildNodes.Remove(rc);
                            rc.Parent = null;
                        }
                    }
                }
            }
            else if(arrData.Count == 0)
            {
                foreach(var el in template.TemplateElements)
                {
                    el.Remove();
                }    
            }    
        }

        private RepeatingTemplate GetRepeatingTemplate(RenderContext context)
        {
            RepeatingTemplate mfTemplate = new RepeatingTemplate();

            //find the ascendant of both start and end bookmark node
            var mfStart = context.MergeField.StartField.StartNode;
            OpenXmlElement parentNode = mfStart;

            while (parentNode.Parent != null && !parentNode.Parent.Descendants().Any(el => el == context.MergeField.EndField.EndNode))
            {
                parentNode = parentNode.Parent;
            }

            if (parentNode != null)
            {
                OpenXmlElement lastChildNode = null;
                OpenXmlElement startNode = parentNode, endNode = null;

                if (context.Evaluator is TableEvaluator)
                {
                    var tableParentNode = startNode;
                    while (tableParentNode != null)
                    {
                        if (tableParentNode is WP.Table || tableParentNode.Descendants<WP.Table>().Any()
                            || tableParentNode == context.MergeField.EndField.EndNode
                            || tableParentNode.Descendants().Any(el => context.MergeField.EndField.EndNode == el))
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

                        if (curentNode == context.MergeField.EndField.EndNode || curentNode.Descendants().Any(el => context.MergeField.EndField.EndNode == el))
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
                        mfTemplate.LastNode = lastChildNode.NextSibling();
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

        private void FillData(List<RenderContext> renderContexts, JObject data)
        {
            foreach (var context in renderContexts)
            {
                var value = data.GetValue(context.FieldName, StringComparison.OrdinalIgnoreCase);

                if (context.Evaluator is LoopEvaluator)
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
                else if (context.Evaluator is ConditionEvaluator)
                {
                    if (value is JValue)
                    {
                        var jvalue = ((JValue)value).Value;
                        var listParam = Utils.PaserParametters(context.Parameters);
                        listParam.Insert(0, context.Operator);
                        var eval = context.Evaluator.Evaluate(jvalue, listParam);
                        if (string.Compare(true.ToString(), eval, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            context.MergeField.StartField?.RemoveAll(true);
                            context.MergeField.EndField?.RemoveAll(true);
                        }
                        else
                        {
                            var start = context.MergeField.StartField.GetAllElements()[0];
                            var end = context.MergeField.EndField.GetAllElements(true)[0];
                            RemoveFromNodeToNode(start, end);                            
                        }
                    }
                }
                else if (context.Evaluator is ImageEvaluator)
                {
                    if (value is JValue)
                    {
                        var jvalue = ((JValue)value).Value;
                        if (jvalue != null)
                        {
                            var base64Img = context.Evaluator.Evaluate(jvalue.ToString(), null);
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
                else if (context.Evaluator is HtmlEvaluator)
                {
                    bool isRemovedMergeField = false;
                    if (value is JValue)
                    {
                        var jvalue = ((JValue)value).Value;
                        var eval = context.Evaluator.Evaluate(jvalue, null);
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
                else if (context.Evaluator is WordEvaluator)
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
                                var eval = context.Evaluator.Evaluate(jvalue, null);
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
                else
                {
                    if (value is JValue)
                    {
                        var jvalue = ((JValue)value).Value;
                        if (jvalue != null)
                        {
                            string eval = "";
                            if (context.Evaluator != null)
                            {
                                try
                                {
                                    eval = context.Evaluator.Evaluate(jvalue, context.Evaluator is DefaultEvaluator ? new List<object>() { context.Parameters } : Utils.PaserParametters(context.Parameters));
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
        }

        private void RemoveFromNodeToNode(OpenXmlElement start, OpenXmlElement end)
        {
            OpenXmlElement lca = null, temp = null;
            bool isFirstParent = false;
            var currentNode = start;
            while (currentNode != null)
            {
                if (currentNode == end || currentNode.Descendants().Any(n => n == end))
                {
                    lca = currentNode;
                    break;
                }

                if (currentNode.InnerText.Trim() == string.Empty)
                {
                    temp = currentNode.NextSibling();
                    if (temp != null)
                    {
                        currentNode.Remove();
                        currentNode = temp;
                    }
                    else
                    {
                        temp = currentNode.Parent;
                        currentNode.Remove();
                        currentNode = temp;
                        isFirstParent = true;
                        continue;
                    }
                }  

                while (currentNode != null)
                {
                    if (currentNode == end || currentNode.Descendants().Any(n => n == end))
                    {
                        lca = currentNode;
                        break;
                    }

                    temp = currentNode.NextSibling();
                    if (temp != null)
                    {
                        if(!isFirstParent)
                            currentNode.Remove();
                        currentNode = temp;
                    }
                    else
                    {
                        temp = currentNode.Parent;
                        if (!isFirstParent)
                            currentNode.Remove();
                        currentNode = temp;
                        isFirstParent = true;
                        break;
                    }
                    isFirstParent = false;
                }                
            }

            if (lca == null) return;

            if (lca == end)
            {
                end.Remove();
            }
            else
            {
                currentNode = end;
                isFirstParent = false;
                while (currentNode != lca)
                {
                    if (currentNode.InnerText.Trim() == string.Empty)
                    {
                        temp = currentNode.PreviousSibling();
                        if (temp != null)
                        {
                            currentNode.Remove();
                            currentNode = temp;
                        }
                        else
                        {
                            temp = currentNode.Parent;
                            currentNode.Remove();
                            currentNode = temp;
                            continue;
                        }
                    }

                    while (currentNode != null)
                    {
                        temp = currentNode.PreviousSibling();
                        if (temp != null)
                        {
                            if(!isFirstParent)
                                currentNode.Remove();
                            currentNode = temp;
                        }
                        else
                        {
                            temp = currentNode.Parent;
                            if (!isFirstParent)
                                currentNode.Remove();
                            currentNode = temp;
                            break;
                        }
                    }
                }

                if (lca.InnerText.Trim() == string.Empty)
                {
                    lca.Remove();
                }
            }    
        }

        private void RemoveFallbackElements(WordprocessingDocument document)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            DocumentFormat.OpenXml.Wordprocessing.Document documentPart = mainPart.Document;

            foreach (var alternative in documentPart.Body.Descendants<AlternateContent>())
            {
                var choice = alternative.Descendants<AlternateContentChoice>().FirstOrDefault();
                if (choice != null)
                {
                    var clonedNodes = choice.ChildElements.Select(x => x.CloneNode(true)).ToList();
                    clonedNodes.ForEach(node => alternative.InsertBeforeSelf(node));
                    alternative.Remove();
                }
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts)
            {
                foreach (var alternative in headerPart.Header.Descendants<AlternateContent>())
                {
                    var choice = alternative.Descendants<AlternateContentChoice>().FirstOrDefault();
                    if (choice != null)
                    {
                        var clonedNodes = choice.ChildElements.Select(x => x.CloneNode(true)).ToList();
                        clonedNodes.ForEach(node => alternative.InsertBeforeSelf(node));
                        alternative.Remove();
                    }
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts)
            {
                foreach (var alternative in footerPart.Footer.Descendants<AlternateContent>())
                {
                    var choice = alternative.Descendants<AlternateContentChoice>().FirstOrDefault();
                    if (choice != null)
                    {
                        var clonedNodes = choice.ChildElements.Select(x => x.CloneNode(true)).ToList();
                        clonedNodes.ForEach(node => alternative.InsertBeforeSelf(node));
                        alternative.Remove();
                    }
                }
            }
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

        public void Dispose()
        {
            if (_sourceStream != null)
                _sourceStream.Dispose();
        }
    }
}
