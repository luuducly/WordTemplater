using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using DRAW = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordTemplater
{
    /// <summary>
    /// Provides functionality for processing Word document templates with data.
    /// </summary>
    public class WordTemplate : IDisposable
    {
        private List<RenderContext> _renderContexts;
        private Stream _sourceStream;
        private Dictionary<string, IEvaluator> _evaluatorFactory;

        /// <summary>
        /// Initializes a new instance of the WordTemplate class.
        /// </summary>
        /// <param name="sourceStream">The template Word file stream.</param>
        /// <exception cref="ArgumentNullException">Thrown when sourceStream is null.</exception>
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
        /// Registers a new format evaluator for template processing.
        /// </summary>
        /// <param name="name">The name of the format evaluator.</param>
        /// <param name="evaluator">An instance of IEvaluator to register.</param>
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
        /// Fills data into the template file and exports it.
        /// </summary>
        /// <param name="data">The input data object (e.g., JObject or any data model).</param>
        /// <param name="removeFallBack">Whether to remove fallback elements after exporting. Default is true.</param>
        /// <returns>The exported file stream.</returns>
        public Stream Export(object data, bool removeFallBack = true)
        {
            if (data != null)
            {
                Stream targetStream = Utils.CloneStream(_sourceStream);
                if (targetStream != null)
                {
                    using (
                        WordprocessingDocument targetDocument = WordprocessingDocument.Open(
                            targetStream,
                            true
                        )
                    )
                    {
                        if (removeFallBack)
                            RemoveFallbackElements(targetDocument);
                        if (data is not JObject)
                            data = JObject.FromObject(data);

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

        /// <summary>
        /// Prepares the render context for the document.
        /// </summary>
        /// <param name="document">The Word document to process.</param>
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

        /// <summary>
        /// Prepares render context for a single element.
        /// </summary>
        /// <param name="element">The OpenXML element to process.</param>
        /// <param name="parentPart">The parent part containing the element.</param>
        /// <returns>A list of render contexts.</returns>
        private List<RenderContext> PrepareRenderContext(
            OpenXmlElement element,
            TypedOpenXmlPart parentPart
        )
        {
            return PrepareRenderContext(new List<OpenXmlElement>() { element }, parentPart);
        }

        /// <summary>
        /// Prepares render context for a list of elements.
        /// </summary>
        /// <param name="elements">The list of OpenXML elements to process.</param>
        /// <param name="parentPart">The parent part containing the elements.</param>
        /// <returns>A list of render contexts.</returns>
        private List<RenderContext> PrepareRenderContext(
            List<OpenXmlElement> elements,
            TypedOpenXmlPart parentPart
        )
        {
            List<RenderContext> rcRootList = new List<RenderContext>();
            Stack<RenderContext> parents = new Stack<RenderContext>();
            Stack<RenderContext> openContexts = new Stack<RenderContext>();
            List<OpenXmlElement> allMergeFieldNodes = new List<OpenXmlElement>();

            foreach (var element in elements)
            {
                allMergeFieldNodes.AddRange(
                    element.Descendants().Where(x => IsMergeFieldNode(x)).ToList()
                );
            }

            for (int i = 0; i < allMergeFieldNodes.Count; i++)
            {
                var mergeFieldNode = allMergeFieldNodes[i];
                var code = GetCode(mergeFieldNode);

                if (code.Contains(FunctionName.EndIf, StringComparison.OrdinalIgnoreCase))
                {
                    if (openContexts.Count > 0)
                        openContexts.Pop().MergeField.EndField = new MergeFieldTemplate(
                            mergeFieldNode
                        );
                    continue;
                }

                if (
                    code.Contains(FunctionName.EndLoop, StringComparison.OrdinalIgnoreCase)
                    || code.Contains(FunctionName.EndTable, StringComparison.OrdinalIgnoreCase)
                )
                {
                    if (parents.Count > 0)
                        parents.Pop();
                    if (openContexts.Count > 0)
                        openContexts.Pop().MergeField.EndField = new MergeFieldTemplate(
                            mergeFieldNode
                        );
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
                if (
                    context.Evaluator != null
                    && (
                        context.Evaluator is LoopEvaluator
                        || context.Evaluator is ConditionEvaluator
                    )
                )
                {
                    if (context.Evaluator is LoopEvaluator)
                        parents.Push(context);

                    openContexts.Push(context);
                }
            }
            return rcRootList;
        }

        /// <summary>
        /// Determines if an element is a merge field node.
        /// </summary>
        /// <param name="x">The element to check.</param>
        /// <returns>True if the element is a merge field node.</returns>
        private bool IsMergeFieldNode(OpenXmlElement x)
        {
            if (x is SimpleField)
            {
                var simpleField = (SimpleField)x;
                var fieldCode = simpleField.Instruction.Value;
                if (!string.IsNullOrEmpty(fieldCode) && fieldCode.Contains(Constant.MERGEFORMAT))
                    return true;
            }
            else if (x is FieldCode)
            {
                var fieldCode = ((FieldCode)x).Text;
                if (!string.IsNullOrEmpty(fieldCode) && fieldCode.Contains(Constant.MERGEFORMAT))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Gets the field code from a node.
        /// </summary>
        /// <param name="node">The node to get the code from.</param>
        /// <returns>The field code as a string.</returns>
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

        /// <summary>
        /// Gets the format template from a code string and sets it in the context.
        /// </summary>
        /// <param name="code">The code string to parse.</param>
        /// <param name="context">The render context to update.</param>
        private void GetFormatTemplate(string code, RenderContext context)
        {
            if (string.IsNullOrEmpty(code))
                return;
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
                if (i3 >= 0)
                {
                    var functionName = code.Substring(0, i3).Trim();
                    var parameters = code.Substring(i3 + 1, code.Length - i3 - 2).Trim();
                    context.FieldName = functionName;
                    context.Parameters = parameters;
                    if (_evaluatorFactory.ContainsKey(functionName.ToLower()))
                    {
                        context.Evaluator = _evaluatorFactory[functionName.ToLower()];
                    }
                }
            }
        }

        /// <summary>
        /// Renders the template with the provided data.
        /// </summary>
        /// <param name="renderContexts">The list of render contexts to process.</param>
        /// <param name="dataObj">The data object to use for rendering.</param>
        private void RenderTemplate(List<RenderContext> renderContexts, JObject dataObj)
        {
            foreach (var rc in renderContexts)
            {
                if (rc.Evaluator is LoopEvaluator)
                {
                    var data = dataObj[rc.FieldName];
                    if (data is JArray arrData)
                    {
                        RenderTemplate(rc, arrData);
                    }
                }
            }
        }

        /// <summary>
        /// Renders a template with array data.
        /// </summary>
        /// <param name="rc">The render context to process.</param>
        /// <param name="arrData">The array data to use for rendering.</param>
        private void RenderTemplate(RenderContext rc, JArray arrData)
        {
            var template = GetRepeatingTemplate(rc);
            if (template != null)
            {
                foreach (var item in arrData)
                {
                    var cloneElements = template.CloneAndAppendTemplate();
                    if (rc.ChildNodes.Count > 0)
                    {
                        foreach (var child in rc.ChildNodes)
                        {
                            if (child.Evaluator is LoopEvaluator)
                            {
                                var childData = item[child.FieldName];
                                if (childData is JArray childArrData)
                                {
                                    RenderTemplate(child, childArrData);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets a repeating template from a render context.
        /// </summary>
        /// <param name="context">The render context to process.</param>
        /// <returns>A repeating template instance.</returns>
        private RepeatingTemplate GetRepeatingTemplate(RenderContext context)
        {
            var template = new RepeatingTemplate();
            var startField = context.MergeField.StartField;
            var endField = context.MergeField.EndField;

            if (startField != null && endField != null)
            {
                var startNode = startField.StartNode;
                var endNode = endField.EndNode;

                if (startNode != null && endNode != null)
                {
                    var parent = startNode.Parent;
                    if (parent != null)
                    {
                        template.ParentNode = parent;
                        var current = startNode;
                        while (current != null && current != endNode)
                        {
                            template.TemplateElements.Add(current);
                            current = current.NextSibling();
                        }
                        if (current != null)
                        {
                            template.TemplateElements.Add(current);
                        }
                        template.LastNode = endNode;
                    }
                }
            }

            return template;
        }

        /// <summary>
        /// Moves a merge field template to a table cell.
        /// </summary>
        /// <param name="fieldTemplate">The merge field template to move.</param>
        /// <param name="tableCell">The target table cell.</param>
        /// <param name="forBeginning">Whether to move to the beginning of the cell.</param>
        private void MoveMergeFieldTo(
            MergeFieldTemplate fieldTemplate,
            TableCell? tableCell,
            bool forBeginning = true
        )
        {
            if (fieldTemplate != null && tableCell != null)
            {
                var elements = fieldTemplate.GetAllElements();
                foreach (var element in elements)
                {
                    if (forBeginning)
                    {
                        tableCell.InsertBefore(element, tableCell.FirstChild);
                    }
                    else
                    {
                        tableCell.Append(element);
                    }
                }
            }
        }

        /// <summary>
        /// Finds a render context in a list.
        /// </summary>
        /// <param name="renderContexts">The list of render contexts to search in.</param>
        /// <param name="renderContext">The render context to find.</param>
        /// <returns>The found render context, or null if not found.</returns>
        private RenderContext Find(List<RenderContext> renderContexts, RenderContext renderContext)
        {
            foreach (var rc in renderContexts)
            {
                if (rc == renderContext)
                    return rc;
                var found = Find(rc.ChildNodes, renderContext);
                if (found != null)
                    return found;
            }
            return null;
        }

        /// <summary>
        /// Fills data into the template.
        /// </summary>
        /// <param name="renderContexts">The list of render contexts to process.</param>
        /// <param name="data">The data object to use for filling.</param>
        private void FillData(List<RenderContext> renderContexts, JObject data)
        {
            foreach (var rc in renderContexts)
            {
                if (rc.Evaluator != null)
                {
                    var value = data[rc.FieldName];
                    if (value != null)
                    {
                        var parameters = Utils.PaserParametters(rc.Parameters);
                        var result = rc.Evaluator.Evaluate(value, parameters);
                        if (rc.MergeField.StartField != null)
                        {
                            var textNode = rc.MergeField.StartField.RemoveAllExceptTextNode();
                            if (textNode != null)
                            {
                                textNode.Text = result;
                            }
                        }
                    }
                }

                if (rc.ChildNodes.Count > 0)
                {
                    FillData(rc.ChildNodes, data);
                }
            }
        }

        /// <summary>
        /// Removes fallback elements from the document.
        /// </summary>
        /// <param name="document">The document to process.</param>
        private void RemoveFallbackElements(WordprocessingDocument document)
        {
            var mainPart = document.MainDocumentPart;
            var documentPart = mainPart.Document;

            // Remove fallback elements from header parts
            foreach (HeaderPart headerPart in mainPart.HeaderParts)
            {
                RemoveFallbackElements(headerPart.Header);
            }

            // Remove fallback elements from body
            RemoveFallbackElements(documentPart.Body);

            // Remove fallback elements from footer parts
            foreach (FooterPart footerPart in mainPart.FooterParts)
            {
                RemoveFallbackElements(footerPart.Footer);
            }
        }

        /// <summary>
        /// Adds an image part to a parent part.
        /// </summary>
        /// <param name="parentPart">The parent part to add the image to.</param>
        /// <returns>The created image part.</returns>
        private ImagePart AddImagePart(TypedOpenXmlPart parentPart)
        {
            return parentPart.AddNewPart<ImagePart>("image/png", "rId" + Utils.GetUintId());
        }

        /// <summary>
        /// Creates a new drawing element for an image.
        /// </summary>
        /// <param name="image">The image element to create a drawing for.</param>
        /// <param name="size">The size of the image.</param>
        /// <returns>The created drawing element.</returns>
        private OpenXmlElement CreateNewDrawingElement(OpenXmlElement image, Size size)
        {
            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = size.Width, Cy = size.Height },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = Utils.GetUintId(),
                        Name = Utils.GetUniqueStringID()
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }
                    ),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = Utils.GetUintId(),
                                        Name = Utils.GetUniqueStringID()
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension()
                                            {
                                                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                            }
                                        )
                                    )
                                    {
                                        Embed = "rId" + Utils.GetUintId(),
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(new A.FillRectangle())
                                ),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = size.Width, Cy = size.Height }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle
                                    }
                                )
                            )
                        )
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    )
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                    RecentPositioning = true
                }
            );

            return drawing;
        }

        /// <summary>
        /// Creates a new picture element.
        /// </summary>
        /// <param name="fileName">The name of the image file.</param>
        /// <param name="width">The width of the image.</param>
        /// <param name="height">The height of the image.</param>
        /// <returns>The created picture element.</returns>
        private PIC.Picture CreateNewPictureElement(string fileName, long width, long height)
        {
            return new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties()
                    {
                        Id = Utils.GetUintId(),
                        Name = Utils.GetUniqueStringID()
                    },
                    new PIC.NonVisualPictureDrawingProperties()
                ),
                new PIC.BlipFill(
                    new A.Blip(
                        new A.BlipExtensionList(
                            new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                        )
                    )
                    {
                        Embed = "rId" + Utils.GetUintId(),
                        CompressionState = A.BlipCompressionValues.Print
                    },
                    new A.Stretch(new A.FillRectangle())
                ),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = 0L, Y = 0L },
                        new A.Extents() { Cx = width, Cy = height }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList())
                    {
                        Preset = A.ShapeTypeValues.Rectangle
                    }
                )
            );
        }

        /// <summary>
        /// Updates the image ID and size of an element.
        /// </summary>
        /// <param name="element">The element to update.</param>
        /// <param name="imageId">The new image ID.</param>
        /// <param name="size">The new size.</param>
        private void UpdateImageIdAndSize(OpenXmlElement element, string imageId, Size size)
        {
            var blip = element.Descendants<A.Blip>().FirstOrDefault();
            if (blip != null)
            {
                blip.Embed = imageId;
            }

            var extents = element.Descendants<DRAW.Extents>().FirstOrDefault();
            if (extents != null)
            {
                extents.Cx = size.Width;
                extents.Cy = size.Height;
            }
        }

        /// <summary>
        /// Gets the size of a shape from its extents.
        /// </summary>
        /// <param name="extents">The extents of the shape.</param>
        /// <returns>The size of the shape.</returns>
        private Size GetShapeSize(DRAW.Extents extents)
        {
            return new Size(extents.Cx, extents.Cy);
        }

        /// <summary>
        /// Disposes of the WordTemplate instance.
        /// </summary>
        public void Dispose()
        {
            if (_sourceStream != null)
            {
                _sourceStream.Dispose();
            }
        }
    }
}
