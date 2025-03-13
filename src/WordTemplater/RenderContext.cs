using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordTemplater
{
    internal class RenderContext
    {
        internal string FieldName { get; set; }
        internal MergeField MergeField { get; set; }
        internal ITemplater Templater { get; set; }
        internal string Parameters { get; set; }
        internal string Operator { get; set; }
        internal RenderContext Parent { get; set; }
        internal int Index { get; set; }
        internal List<RenderContext> ChildNodes { get; set; }


        public RenderContext()
        {
            ChildNodes = new List<RenderContext>();
        }

        public override string ToString()
        {
            return FieldName + ":" + Templater.ToString() + ":" + Parameters;
        }
    }

    internal class MergeField
    {
        internal MergeFieldTemplate StartField { get; set; }
        internal MergeFieldTemplate EndField { get; set; }
        internal TypedOpenXmlPart ParentPart { get; set; }
    }

    internal class MergeFieldTemplate
    {
        internal OpenXmlElement StartNode
        {
            get
            {
                if (_simpleField != null) return _simpleField;
                return _beginFieldChar;
            }
        }

        internal OpenXmlElement EndNode
        {
            get
            {
                if (_simpleField != null) return _simpleField;
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

        internal void RemoveAll(bool deleteParent = false)
        {
            if (_isRemoved) return;
            var parentParagraph = StartNode.Ancestors<Paragraph>().FirstOrDefault();
            foreach (var el in _allElements)
            {
                if(el.Parent != null)
                    el.Remove();
            }
            if (deleteParent && parentParagraph != null && parentParagraph.Parent != null && parentParagraph.InnerText.Trim() == string.Empty)
            {
                if(parentParagraph.Parent.ChildElements.Count(c => c is Paragraph) > 1)
                    parentParagraph.Remove();
            }
            _isRemoved = true;
        }

        internal WP.Text RemoveAllExceptTextNode()
        {
            if (_isRemoved) return _textNode;
            if (_simpleField != null)
            {
                var run = _simpleField.Descendants<Run>().FirstOrDefault();
                if(run != null)
                {
                    _textNode = run.Descendants<WP.Text>().FirstOrDefault();
                    if(run.Parent != null)
                        run.Remove();
                    if (_simpleField.Parent != null)
                    {
                        _simpleField.InsertBeforeSelf(run);
                        _simpleField.Remove();
                    }
                }    
            }
            else
            {
                foreach (var el in _allElements)
                {
                        if (_textNode == null)
                        {
                            _textNode = el.Descendants<WP.Text>().FirstOrDefault();
                            if (_textNode == null && el.Parent != null) el.Remove();
                        }
                        else if(el.Parent != null)
                            el.Remove();
                }
            }
            _isRemoved = true;
            return _textNode;
        }

        private FieldChar FindFieldChar(FieldCode fieldCode, FieldCharValues type)
        {
            if (fieldCode == null) return null;
            var parent = fieldCode.Parent;
            while(parent != null)
            {
                if (!_allElements.Contains(parent))
                {
                    if (type == FieldCharValues.End)
                        _allElements.Add(parent);
                    else
                        _allElements.Insert(0, parent);

                }
                var fieldChar = parent.Descendants<FieldChar>().Where(fc => fc.FieldCharType == type).FirstOrDefault();
                if(fieldChar != null) return fieldChar;
                if (type == FieldCharValues.End)
                    parent = parent.NextSibling();
                else
                    parent = parent.PreviousSibling();
            }
            return null;
        }
    }

    internal class RepeatingTemplate
    {
        internal OpenXmlElement LastNode;
        internal OpenXmlElement ParentNode;
        internal List<OpenXmlElement> TemplateElements;

        internal RepeatingTemplate()
        {
            TemplateElements = new List<OpenXmlElement>();
        }

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

        private List<OpenXmlElement> CloneTemplateElements()
        {
            List<OpenXmlElement> templateElements = new List<OpenXmlElement>();
            foreach (OpenXmlElement templateElement in TemplateElements)
            {
                templateElements.Add(templateElement.CloneNode(true));
            }
            return templateElements;
        }

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
