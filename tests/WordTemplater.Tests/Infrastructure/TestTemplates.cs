using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace WordTemplater.Tests
{
    /// <summary>
    /// Programmatic templates and data builders shared by the golden, lifetime, concurrency and mutation tests.
    /// </summary>
    internal static class TestTemplates
    {
        /// <summary>
        /// A merge field the library recognizes (the code must contain "  \* MERGEFORMAT" with two spaces).
        /// </summary>
        public static SimpleField F(string code) =>
            new SimpleField(new Run(new Text("«" + code + "»"))) { Instruction = " MERGEFIELD  " + code + "  \\* MERGEFORMAT " };

        public static Paragraph P(params OpenXmlElement[] children) => new Paragraph(children);

        public static Run T(string text) => new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        /// <summary>
        /// Builds a .docx whose body contains the given top-level elements.
        /// </summary>
        public static byte[] Build(params OpenXmlElement[] bodyChildren)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(bodyChildren));
            }
            return ms.ToArray();
        }

        /// <summary>
        /// Title, then table(Items) over a table with a plain header row and one template row, then endtable.
        /// </summary>
        public static byte[] TableTemplate(string? extraCellCode = null)
        {
            var header = new TableRow(
                new TableCell(P(T("No"))),
                new TableCell(P(T("Name"))),
                new TableCell(P(T("Qty"))),
                new TableCell(P(T("Price"))));
            var row = new TableRow(
                new TableCell(P(F("_index"))),
                new TableCell(P(F("Name:upper()"))),
                new TableCell(P(F("Qty:number(#,##0)"))),
                new TableCell(P(F("Price:currency(en-US)"))));
            if (extraCellCode != null)
            {
                row.Append(new TableCell(P(F(extraCellCode))));
            }
            return Build(
                P(F("Title")),
                P(F("table(Items)")),
                new Table(header, row),
                P(F("endtable")),
                P(T("end")));
        }

        /// <summary>
        /// Like <see cref="TableTemplate"/>, but the table also has a plain-text footer row after the template row,
        /// so the repeated rows are inserted before an existing row.
        /// </summary>
        public static byte[] TableWithFooterTemplate()
        {
            var header = new TableRow(
                new TableCell(P(T("No"))),
                new TableCell(P(T("Name"))),
                new TableCell(P(T("Qty"))),
                new TableCell(P(T("Price"))));
            var row = new TableRow(
                new TableCell(P(F("_index"))),
                new TableCell(P(F("Name:upper()"))),
                new TableCell(P(F("Qty:number(#,##0)"))),
                new TableCell(P(F("Price:currency(en-US)"))));
            var footer = new TableRow(
                new TableCell(P(T("Total"))),
                new TableCell(P()),
                new TableCell(P()),
                new TableCell(P()));
            return Build(
                P(F("Title")),
                P(F("table(Items)")),
                new Table(header, row, footer),
                P(F("endtable")),
                P(T("end")));
        }

        /// <summary>
        /// Paragraph loop over Groups containing a paragraph loop over Items with a nested if(Qty &gt; 5) block.
        /// </summary>
        public static byte[] NestedLoopTemplate()
        {
            return Build(
                P(F("Title")),
                P(F("loop(Groups)")),
                P(F("Name")),
                P(F("loop(Items)")),
                P(F("_index"), T(". "), F("Name")),
                P(F("if(Qty > 5)")),
                P(T("big: "), F("Qty")),
                P(F("endif")),
                P(F("endloop")),
                P(T("after inner")),
                P(F("endloop")),
                P(T("end")));
        }

        /// <summary>
        /// Paragraph loop over Items with a nested if(Qty &gt; 5) block.
        /// </summary>
        public static byte[] LoopIfTemplate()
        {
            return Build(
                P(F("Title")),
                P(F("loop(Items)")),
                P(F("_index"), T(". "), F("Name")),
                P(F("if(Qty > 5)")),
                P(T("big qty: "), F("Qty")),
                P(F("endif")),
                P(F("endloop")),
                P(T("end")));
        }

        /// <summary>
        /// Paragraph loop over Items rendering the loop metadata fields separated by "|".
        /// </summary>
        public static byte[] LoopMetadataTemplate()
        {
            return Build(
                P(F("loop(Items)")),
                P(F("_index"), T("|"), F("_first"), T("|"), F("_last"), T("|"), F("_even"), T("|"), F("_odd"), T("|"), F("Name")),
                P(F("endloop")));
        }

        public static JObject Data(int n, bool withCode = false)
        {
            var items = new JArray();
            for (int i = 0; i < n; i++)
            {
                var item = new JObject
                {
                    ["Name"] = "item " + i,
                    ["Qty"] = i % 10,
                    ["Price"] = i * 1.5m
                };
                if (withCode)
                {
                    item["Code"] = "CODE-" + i;
                }
                items.Add(item);
            }
            return new JObject
            {
                ["Title"] = "Report",
                ["Items"] = items
            };
        }

        /// <summary>
        /// Data for <see cref="NestedLoopTemplate"/>: Groups[g] = { Name, Items[itemsPerGroup] }; group 1 (if any) has no items.
        /// </summary>
        public static JObject NestedData(int groups, int itemsPerGroup)
        {
            var groupArray = new JArray();
            for (int g = 0; g < groups; g++)
            {
                var items = new JArray();
                if (g != 1)
                {
                    for (int i = 0; i < itemsPerGroup; i++)
                    {
                        var n = g * itemsPerGroup + i;
                        items.Add(new JObject
                        {
                            ["Name"] = "item " + g + "." + i,
                            ["Qty"] = n % 10,
                            ["Price"] = n * 1.5m
                        });
                    }
                }
                groupArray.Add(new JObject
                {
                    ["Name"] = "group " + g,
                    ["Items"] = items
                });
            }
            return new JObject
            {
                ["Title"] = "Report",
                ["Groups"] = groupArray
            };
        }
        /// <summary>
        /// Creates a modern DrawingML Text Shape anchored with relative positioning for preprinted bill forms.
        /// </summary>
        public static Run CreateDrawingMLTextShape(string fieldCode, double leftPt, double topPt, double widthPt, double heightPt, string fontColor = "0F172A", string fontSizeHalfPt = "22", bool bold = false, bool italic = false)
        {
            long xEmu = (long)(leftPt * 12700);
            long yEmu = (long)(topPt * 12700);
            long wEmu = (long)(widthPt * 12700);
            long hEmu = (long)(heightPt * 12700);

            string boldXml = bold ? "<w:b/>" : "";
            string italicXml = italic ? "<w:i/>" : "";
            string colorXml = !string.IsNullOrEmpty(fontColor) ? $"<w:color w:val=\"{fontColor}\"/>" : "";
            string szXml = !string.IsNullOrEmpty(fontSizeHalfPt) ? $"<w:sz w:val=\"{fontSizeHalfPt}\"/>" : "";

            uint shapeId = (uint)(System.Guid.NewGuid().GetHashCode() & 0x7FFFFFFF);

            string drawingInnerXml = $@"<wp:anchor xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
                        xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
                        xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""
                        xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
                        distT=""0"" distB=""0"" distL=""0"" distR=""0"" simplePos=""0"" relativeHeight=""251658240"" behindDoc=""0"" locked=""0"" layoutInCell=""1"" allowOverlap=""1"">
  <wp:simplePos x=""0"" y=""0""/>
  <wp:positionH relativeFrom=""column"">
    <wp:posOffset>{xEmu}</wp:posOffset>
  </wp:positionH>
  <wp:positionV relativeFrom=""paragraph"">
    <wp:posOffset>{yEmu}</wp:posOffset>
  </wp:positionV>
  <wp:extent cx=""{wEmu}"" cy=""{hEmu}""/>
  <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
  <wp:wrapNone/>
  <wp:docPr id=""{shapeId}"" name=""TextShape_{shapeId}""/>
  <wp:cNvGraphicFramePr/>
  <a:graphic>
    <a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp>
        <wps:cNvSpPr txBox=""1""/>
        <wps:spPr>
          <a:xfrm>
            <a:off x=""0"" y=""0""/>
            <a:ext cx=""{wEmu}"" cy=""{hEmu}""/>
          </a:xfrm>
          <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln><a:noFill/></a:ln>
        </wps:spPr>
        <wps:txbx>
          <w:txbxContent>
            <w:p>
              <w:pPr><w:spacing w:before=""0"" w:after=""0"" w:line=""240"" w:lineRule=""auto""/></w:pPr>
              <w:fldSimple w:instr="" MERGEFIELD {fieldCode} \* MERGEFORMAT "">
                <w:r>
                  <w:rPr>{boldXml}{italicXml}{colorXml}{szXml}</w:rPr>
                  <w:t>«{fieldCode}»</w:t>
                </w:r>
              </w:fldSimple>
            </w:p>
          </w:txbxContent>
        </wps:txbx>
        <wps:bodyPr rot=""0"" spcFirstLastPara=""0"" vertOverflow=""overflow"" horzOverflow=""overflow"" vert=""horz"" wrap=""none"" lIns=""0"" tIns=""0"" rIns=""0"" bIns=""0"" numCol=""1"" spcCol=""0"">
          <a:noFill/>
        </wps:bodyPr>
      </wps:wsp>
    </a:graphicData>
  </a:graphic>
</wp:anchor>";

            var drawing = new Drawing { InnerXml = drawingInnerXml };
            return new Run(drawing);
        }

        public static byte[] BuildPreprintedBillTemplate()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document();

                // 1. Header with exact text shapes over pre-printed dots
                var headerPart = main.AddNewPart<HeaderPart>();
                headerPart.Header = new Header(
                    P(T("CÔNG TY TNHH PHÁT TRIỂN CÔNG NGHỆ MẪU (EXAMPLE CORP)")),
                    P(T("MST: 0100000000-001  |  ĐC: Số 1 Đường Thí Nghiệm, Phường Demo, TP. Hà Nội")),
                    P(
                        T("Mẫu số: ..................       Ký hiệu: ..................       Số: .................."),
                        CreateDrawingMLTextShape("InvoiceForm", leftPt: 45, topPt: -1, widthPt: 85, heightPt: 16, fontColor: "1E293B", bold: true),
                        CreateDrawingMLTextShape("InvoiceSerial", leftPt: 175, topPt: -1, widthPt: 75, heightPt: 16, fontColor: "1E293B", bold: true),
                        CreateDrawingMLTextShape("padleft(InvoiceNo, 7, '0')", leftPt: 295, topPt: -2, widthPt: 90, heightPt: 18, fontColor: "DC2626", fontSizeHalfPt: "26", bold: true)
                    ),
                    P(
                        T("Ngày lập hóa đơn: Ngày ..... tháng ..... năm 2026"),
                        CreateDrawingMLTextShape("BillDay", leftPt: 118, topPt: -1, widthPt: 25, heightPt: 16, fontColor: "1E293B", bold: true),
                        CreateDrawingMLTextShape("BillMonth", leftPt: 178, topPt: -1, widthPt: 25, heightPt: 16, fontColor: "1E293B", bold: true)
                    )
                );

                // 2. Footer with exact text shapes for signatures
                var footerPart = main.AddNewPart<FooterPart>();
                footerPart.Footer = new Footer(
                    P(T("Người mua hàng                     Người giao hàng                     Thủ trưởng đơn vị")),
                    P(
                        T("............................        ............................        ............................"),
                        CreateDrawingMLTextShape("BuyerSign", leftPt: 10, topPt: 15, widthPt: 130, heightPt: 18, fontColor: "1E293B", bold: true),
                        CreateDrawingMLTextShape("DelivererSign", leftPt: 175, topPt: 15, widthPt: 130, heightPt: 18, fontColor: "1E293B", bold: true),
                        CreateDrawingMLTextShape("ManagerSign:upper()", leftPt: 345, topPt: 15, widthPt: 140, heightPt: 18, fontColor: "1E293B", bold: true)
                    )
                );

                var sectionProps = new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footerPart) },
                    new PageSize { Width = 11906, Height = 16838 },
                    new PageMargin { Top = 1000, Bottom = 1000, Left = 1000, Right = 1000, Header = 500, Footer = 500 }
                );

                // 3. Body with dotted background lines & overlay Text Shapes
                var body = new Body(
                    P(T("HÓA ĐƠN BÁN HÀNG / GIÁ TRỊ GIA TĂNG")),
                    P(
                        T("Họ tên người mua hàng: .........................................."),
                        CreateDrawingMLTextShape("CustomerName", leftPt: 125, topPt: -1, widthPt: 320, heightPt: 18, fontColor: "0F172A", bold: true)
                    ),
                    P(
                        T("Tên đơn vị (Doanh nghiệp): ...................................."),
                        CreateDrawingMLTextShape("CustomerCompany", leftPt: 140, topPt: -1, widthPt: 320, heightPt: 18, fontColor: "0F172A", bold: false)
                    ),
                    P(
                        T("Mã số thuế: ....................................................."),
                        CreateDrawingMLTextShape("CustomerTaxCode", leftPt: 70, topPt: -1, widthPt: 180, heightPt: 18, fontColor: "0F172A", bold: true)
                    ),
                    P(
                        T("Địa chỉ trụ sở / giao hàng: ...................................."),
                        CreateDrawingMLTextShape("CustomerAddress", leftPt: 135, topPt: -1, widthPt: 330, heightPt: 18, fontColor: "0F172A", bold: false)
                    ),
                    P(
                        T("Hình thức thanh toán: ..............................   Số tài khoản: ...................."),
                        CreateDrawingMLTextShape("PaymentMethod", leftPt: 120, topPt: -1, widthPt: 110, heightPt: 18, fontColor: "0F172A"),
                        CreateDrawingMLTextShape("BankAccount", leftPt: 325, topPt: -1, widthPt: 170, heightPt: 18, fontColor: "0F172A")
                    ),
                    P(),
                    P(
                        T("Cộng tiền hàng: ......................................................................................."),
                        CreateDrawingMLTextShape("Subtotal:currency(vi-VN)", leftPt: 100, topPt: -1, widthPt: 200, heightPt: 18, fontColor: "0F172A", bold: true)
                    ),
                    P(
                        T("Tiền thuế GTGT (10%): ................................................................................."),
                        CreateDrawingMLTextShape("VatAmount:currency(vi-VN)", leftPt: 135, topPt: -1, widthPt: 200, heightPt: 18, fontColor: "0F172A", bold: true)
                    ),
                    P(
                        T("Tổng cộng tiền thanh toán: ....................................................................."),
                        CreateDrawingMLTextShape("TotalAmount:currency(vi-VN)", leftPt: 155, topPt: -2, widthPt: 220, heightPt: 20, fontColor: "DC2626", fontSizeHalfPt: "24", bold: true)
                    ),
                    P(
                        T("Số tiền viết bằng chữ: ............................................................................."),
                        CreateDrawingMLTextShape("TotalInWords", leftPt: 125, topPt: -1, widthPt: 330, heightPt: 18, fontColor: "0F172A", fontSizeHalfPt: "20", bold: true, italic: true)
                    ),
                    sectionProps
                );

                main.Document.Body = body;
            }
            return ms.ToArray();
        }

        private static Run T(string text, bool bold = false, string? color = null, string? size = null)
        {
            var runProps = new RunProperties();
            if (bold) runProps.Append(new Bold());
            if (!string.IsNullOrEmpty(color)) runProps.Append(new Color { Val = color });
            if (!string.IsNullOrEmpty(size)) runProps.Append(new FontSize { Val = size });
            return new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        }

        private static TableCell C(OpenXmlElement content, string? bgColor = null)
        {
            var cell = new TableCell(content);
            if (!string.IsNullOrEmpty(bgColor))
            {
                cell.Append(new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = bgColor }));
            }
            return cell;
        }

        public static byte[] BuildShowcaseTemplate()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();

                var tableProps = new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4, Color = "CBD5E1" },
                        new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "CBD5E1" },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "E2E8F0" },
                        new InsideVerticalBorder { Val = BorderValues.None },
                        new LeftBorder { Val = BorderValues.None },
                        new RightBorder { Val = BorderValues.None }
                    )
                );

                var prodHeader = new TableRow(
                    C(P(T("#", true, "1E293B")), bgColor: "E2E8F0"),
                    C(P(T("SKU", true, "1E293B")), bgColor: "E2E8F0"),
                    C(P(T("Description", true, "1E293B")), bgColor: "E2E8F0"),
                    C(P(T("Qty", true, "1E293B")), bgColor: "E2E8F0"),
                    C(P(T("Unit Price", true, "1E293B")), bgColor: "E2E8F0"),
                    C(P(T("Total", true, "1E293B")), bgColor: "E2E8F0")
                );

                var prodRow = new TableRow(
                    C(P(F("_index"))),
                    C(P(F("Sku:upper()"))),
                    C(P(F("Name"))),
                    C(P(F("Qty:number(#,##0)"))),
                    C(P(F("Price:currency(en-US)"))),
                    C(P(F("LineTotal:currency(en-US)")))
                );

                var body = new Body(
                    // 1. Header Banner & Title
                    P(T("WORDTEMPLATER ENTERPRISE SHOWCASE", true, "0F172A", "36")),
                    P(T("Comprehensive Multi-Data-Type Test Suite for OpenXML Document Automation", false, "64748B", "20")),
                    P(),

                    // 2. Raster Image (Logo)
                    P(T("1. Embedded Raster Image (Base64 PNG):", true, "1E293B", "22")),
                    P(F("image(LogoImage)")),
                    P(),

                    // 3. Document Metadata & String Formatting
                    P(T("2. Document Metadata & String Formatting:", true, "1E293B", "22")),
                    P(T("Document ID (Padded 8 digits): ", true), F("padleft(DocId, 8, '0')"),
                      T("   |   Reference: ", true), F("RefCode:padright(12, '*')")),
                    P(T("Generated At: ", true), F("date(GeneratedAt, 'dddd, dd MMMM yyyy HH:mm', 'en-US')")),
                    P(),

                    // 4. Machine-Readable Identifiers (QR Code & Barcode)
                    P(T("3. Visual Machine-Readable Codes (QR Code & Barcode):", true, "1E293B", "22")),
                    P(T("1D Tracking Barcode (Code128):", true, "334155")),
                    P(F("barcode(BarcodePayload)")),
                    P(T("Barcode Value: ", true), F("BarcodePayload")),
                    P(),
                    P(T("2D Verification QR Code:", true, "334155")),
                    P(F("qrcode(QrPayload)")),
                    P(T("Encoded Target URL: ", true), F("QrPayload")),
                    P(),

                    // 5. Conditional Logic & Dynamic Branches
                    P(T("4. Conditional Logic & Dynamic Branches:", true, "1E293B", "22")),
                    P(T("Verification Status (Inline Boolean): ", true),
                      F("if(IsVerified, 'VERIFIED & CERTIFIED (Digital Signature Valid)', 'UNVERIFIED DRAFT')")),
                    P(T("Performance Incentive (Inline Boolean): ", true),
                      F("if(HasBonus, 'APPROVED - Performance Bonus Granted', 'NOT ELIGIBLE')")),
                    P(T("VIP Tier Assessment (Inline Comparison Score >= 80): ", true),
                      F("if(Score >= 80, 'PLATINUM VIP (Tier 1)', 'STANDARD MEMBER')")),
                    P(),

                    // Page break for clean layout
                    P(new Run(new Break { Type = BreakValues.Page })),

                    // 6. Dynamic Repeating Table with Evaluators & Zebra Striping
                    P(T("5. Dynamic Repeating Table with Evaluators & Zebra Striping:", true, "1E293B", "22")),
                    P(F("table(Products)")),
                    new Table(tableProps, prodHeader, prodRow),
                    P(F("endtable")),
                    P(),

                    // Financial Summary
                    P(T("Subtotal: ", true), F("Subtotal:currency(en-US)"),
                      T("   |   Tax Rate: ", true), F("TaxRate:percentage(2)"),
                      T("   |   Net Total: ", true, "0F172A", "24"), F("Total:currency(en-US)")),
                    P(),

                    // 7. Rich Formatted HTML Content
                    P(T("6. Rich Formatted HTML Content (AltChunk Ingestion):", true, "1E293B", "22")),
                    P(F("html(HtmlNotes)")),
                    P(),

                    // 8. Mermaid Architecture Diagram
                    P(T("7. Mermaid Architecture Diagram (Automated Kroki Service):", true, "1E293B", "22")),
                    P(F("mermaid(ArchitectureDiagram)")),

                    new SectionProperties(
                        new PageSize { Width = 11906, Height = 16838 },
                        new PageMargin { Top = 1000, Bottom = 1000, Left = 1000, Right = 1000 }
                    )
                );

                main.Document = new Document(body);
            }
            return ms.ToArray();
        }
    }
}
