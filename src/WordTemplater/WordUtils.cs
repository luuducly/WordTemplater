using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTemplater
{
  internal static class WordUtils
  {
    internal static void RemoveFromNodeToNode(OpenXmlElement start, OpenXmlElement end)
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
            if (!isFirstParent)
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
              if (!isFirstParent)
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

    internal static void RemoveFallbackElements(WordprocessingDocument document)
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
  }
}
