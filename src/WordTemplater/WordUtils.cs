using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace WordTemplater
{
  internal static class WordUtils
  {
    internal static void RemoveFromNodeToNode(OpenXmlElement start, OpenXmlElement end)
    {
      if (start.Parent != end.Parent)
      {
        start = start.Parent;
        end = end.Parent;
      }
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

    internal static (double width, double height) GetPageSize(OpenXmlElement element)
    {
      Body body = element.Ancestors<Body>().FirstOrDefault();
      if (body == null)
        return (0, 0);

      SectionProperties sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
      if (sectionProps == null) return (0, 0);

      PageSize pageSize = sectionProps.Elements<PageSize>().FirstOrDefault();
      if (pageSize == null) return (0, 0);

      return (TwipToPixels(pageSize.Width), TwipToPixels(pageSize.Height));
    }

    internal static Size GetImageSize(Stream stream)
    {
      stream.Position = 0;
      var image = SKImage.FromEncodedData(stream);
      stream.Position = 0;
      if (image != null)
      {
        var width = image.Width;
        var height = image.Height;
        return new Size(width, height);
      }
      return null;
    }

    internal static Int64Value PixelToEmu(double pixel)
    {
      return (Int64Value)(pixel / Constant.DEFAULT_DPI * Constant.PIXEL_PER_INCH);
    }

    internal static double EmuToPixels(double emu)
    {
      return emu * Constant.DEFAULT_DPI / Constant.PIXEL_PER_INCH;
    }

    internal static double TwipToInches(double twip)
    {
      return twip / Constant.TWIP_PER_INCH;
    }

    internal static double TwipToPixels(double twip)
    {
      return twip / Constant.TWIP_PER_INCH * Constant.DEFAULT_DPI;
    }

    internal static int ToThousandPercent(double percent)
    {
      return (int)(percent * 1000);
    }
  }
}
