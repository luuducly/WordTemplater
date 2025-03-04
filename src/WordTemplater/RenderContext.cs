using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordTemplater
{
    /// <summary>
    /// Represents the context for rendering a template field in a Word document.
    /// </summary>
    internal class RenderContext
    {
        /// <summary>
        /// Gets or sets the name of the field being rendered.
        /// </summary>
        internal string FieldName { get; set; }

        /// <summary>
        /// Gets or sets the merge field information.
        /// </summary>
        internal MergeField MergeField { get; set; }

        /// <summary>
        /// Gets or sets the evaluator used to process the field value.
        /// </summary>
        internal IEvaluator Evaluator { get; set; }

        /// <summary>
        /// Gets or sets the parameters for the field evaluation.
        /// </summary>
        internal string Parameters { get; set; }

        /// <summary>
        /// Gets or sets the operator used in conditional expressions.
        /// </summary>
        internal string Operator { get; set; }

        /// <summary>
        /// Gets or sets the parent context for nested fields.
        /// </summary>
        internal RenderContext Parent { get; set; }

        /// <summary>
        /// Gets or sets the index of the current item in a collection.
        /// </summary>
        internal int Index { get; set; }

        /// <summary>
        /// Gets or sets the list of child contexts for nested fields.
        /// </summary>
        internal List<RenderContext> ChildNodes { get; set; }

        /// <summary>
        /// Initializes a new instance of the RenderContext class.
        /// </summary>
        public RenderContext()
        {
            ChildNodes = new List<RenderContext>();
        }

        /// <summary>
        /// Returns a string representation of the render context.
        /// </summary>
        /// <returns>A string containing the field name, evaluator, and parameters.</returns>
        public override string ToString()
        {
            return FieldName + ":" + Evaluator.ToString() + ":" + Parameters;
        }
    }

    /// <summary>
    /// Represents a merge field in a Word document.
    /// </summary>
    internal class MergeField
    {
        /// <summary>
        /// Gets or sets the start field template.
        /// </summary>
        internal MergeFieldTemplate StartField { get; set; }

        /// <summary>
        /// Gets or sets the end field template.
        /// </summary>
        internal MergeFieldTemplate EndField { get; set; }

        /// <summary>
        /// Gets or sets the parent part containing the merge field.
        /// </summary>
        internal TypedOpenXmlPart ParentPart { get; set; }
    }

    /// <summary>
    /// Represents a template for a merge field in a Word document.
    /// </summary>
    internal class MergeFieldTemplate
    {
        /// <summary>
        /// Gets the start node of the merge field.
        /// </summary>
        internal OpenXmlElement StartNode
        {
            get
            {
                if (_simpleField != null)
                    return _simpleField;
                return _beginFieldChar;
            }
        }

        /// <summary>
        /// Gets the end node of the merge field.
        /// </summary>
        internal OpenXmlElement EndNode
        {
            get
            {
                if (_simpleField != null)
                    return _simpleField;
                return _endFieldChar;
            }
        }

        private FieldChar _beginFieldChar;
        private FieldChar _endFieldChar;
        private FieldCode _fieldCode;
        private SimpleField _simpleField;
        private WP.Text _textNode;
        private List<OpenXmlElement> _allElements;
        private bool _isRemoved = false;

        /// <summary>
        /// Initializes a new instance of the MergeFieldTemplate class.
        /// </summary>
        /// <param name="node">The OpenXML element representing the merge field.</param>
        internal MergeFieldTemplate(OpenXmlElement node)
        {
            _allElements = new List<OpenXmlElement>();
            if (node is FieldCode)
            {
                _fieldCode = (FieldCode)node;
                _beginFieldChar = FindFieldChar(_fieldCode, FieldCharValues.Begin);
                _endFieldChar = FindFieldChar(_fieldCode, FieldCharValues.End);
            }
            else if (node is SimpleField)
            {
                _simpleField = (SimpleField)node;
                _allElements.Add(_simpleField);
            }
        }

        /// <summary>
        /// Gets all elements in the merge field template.
        /// </summary>
        /// <param name="reverse">Whether to return elements in reverse order.</param>
        /// <returns>A list of OpenXML elements.</returns>
        internal List<OpenXmlElement> GetAllElements(bool reverse = false)
        {
            if (!reverse)
            {
                return _allElements.ToList();
            }
            else
            {
                var returnList = _allElements.ToList();
                returnList.Reverse();
                return returnList;
            }
        }

        /// <summary>
        /// Removes all elements from the merge field template.
        /// </summary>
        internal void RemoveAll()
        {
            if (_isRemoved)
                return;
            foreach (var el in _allElements)
            {
                el.Remove();
            }
            _isRemoved = true;
        }

        /// <summary>
        /// Removes all elements except the text node from the merge field template.
        /// </summary>
        /// <returns>The remaining text node.</returns>
        internal WP.Text RemoveAllExceptTextNode()
        {
            if (_isRemoved)
                return _textNode;
            if (_simpleField != null)
            {
                var run = _simpleField.Descendants<Run>().FirstOrDefault();
                if (run != null)
                {
                    _textNode = run.Descendants<WP.Text>().FirstOrDefault();
                    run.Remove();
                    _simpleField.InsertBeforeSelf(run);
                    _simpleField.Remove();
                }
            }
            else
            {
                foreach (var el in _allElements)
                {
                    if (_textNode == null)
                    {
                        _textNode = el.Descendants<WP.Text>().FirstOrDefault();
                        if (_textNode == null)
                            el.Remove();
                    }
                    else
                        el.Remove();
                }
            }
            _isRemoved = true;
            return _textNode;
        }

        /// <summary>
        /// Finds a field character of the specified type in the field code.
        /// </summary>
        /// <param name="fieldCode">The field code to search in.</param>
        /// <param name="type">The type of field character to find.</param>
        /// <returns>The found field character, or null if not found.</returns>
        private FieldChar FindFieldChar(FieldCode fieldCode, FieldCharValues type)
        {
            if (fieldCode == null)
                return null;
            var parent = fieldCode.Parent;
            while (parent != null)
            {
                if (!_allElements.Contains(parent))
                {
                    if (type == FieldCharValues.End)
                        _allElements.Add(parent);
                    else
                        _allElements.Insert(0, parent);
                }
                var fieldChar = parent
                    .Descendants<FieldChar>()
                    .Where(fc => fc.FieldCharType == type)
                    .FirstOrDefault();
                if (fieldChar != null)
                    return fieldChar;
                if (type == FieldCharValues.End)
                    parent = parent.NextSibling();
                else
                    parent = parent.PreviousSibling();
            }
            return null;
        }
    }

    /// <summary>
    /// Represents a template that can be repeated in a Word document.
    /// </summary>
    internal class RepeatingTemplate
    {
        /// <summary>
        /// Gets or sets the last node in the template.
        /// </summary>
        internal OpenXmlElement LastNode;

        /// <summary>
        /// Gets or sets the parent node of the template.
        /// </summary>
        internal OpenXmlElement ParentNode;

        /// <summary>
        /// Gets or sets the list of template elements.
        /// </summary>
        internal List<OpenXmlElement> TemplateElements;

        /// <summary>
        /// Initializes a new instance of the RepeatingTemplate class.
        /// </summary>
        internal RepeatingTemplate()
        {
            TemplateElements = new List<OpenXmlElement>();
        }

        /// <summary>
        /// Clones and appends the template to the document.
        /// </summary>
        /// <returns>A list of the cloned elements.</returns>
        internal List<OpenXmlElement> CloneAndAppendTemplate()
        {
            var cloneElements = CloneTemplateElements();
            if (LastNode != null)
            {
                foreach (var el in cloneElements)
                {
                    GenerateNewIdAndName(el);
                    LastNode.InsertBeforeSelf(el);
                }
            }
            else if (ParentNode != null)
            {
                foreach (var el in cloneElements)
                {
                    GenerateNewIdAndName(el);
                    ParentNode.Append(el);
                }
            }
            return cloneElements;
        }

        /// <summary>
        /// Clones all template elements.
        /// </summary>
        /// <returns>A list of cloned elements.</returns>
        private List<OpenXmlElement> CloneTemplateElements()
        {
            List<OpenXmlElement> templateElements = new List<OpenXmlElement>();
            foreach (OpenXmlElement templateElement in TemplateElements)
            {
                templateElements.Add(templateElement.CloneNode(true));
            }
            return templateElements;
        }

        /// <summary>
        /// Generates new IDs and names for drawing elements.
        /// </summary>
        /// <param name="element">The element to process.</param>
        private void GenerateNewIdAndName(OpenXmlElement element)
        {
            if (element != null)
            {
                foreach (var drawing in element.Descendants<Drawing>())
                {
                    var prop = drawing.Descendants<DocProperties>().FirstOrDefault();
                    if (prop != null)
                    {
                        prop.Id = Utils.GetUintId();
                        prop.Name = Utils.GetUniqueStringID();
                    }
                }
            }
        }
    }
}
