using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXMLCore
{
    public static class OpenXMLExecute
    {
        private static readonly string _imgPath = @"C:\Users\usr\Code\MyProject\OpenXMLDemo\primitive.jpg";
        public static void RunProcess(string filePath)
        {
            if (File.Exists(filePath)) 
            {
                File.Delete(filePath);
            }

            // create docx
            CreateDocx(filePath);

            // add own define style
            AddTitleSytle(filePath);

            AddParagraphText(filePath, "基于模型的建模仿真平台");
            SetParagraphStyle_Title_Level(filePath, "tl1");
            AddParagraphText(filePath, "概述");
            SetParagraphStyle_Title_Level(filePath, "tl2");
            AddParagraphText(filePath, "文档概述");
            SetParagraphStyle_Title_Level(filePath, "tl3");

            string para_1 = "本文档旨对基于模型的建模仿真平台的技术要求响应情况进行描述和说明。本文档为中间版本，非正式文档，后续投标活动中技术方案会基于此进行组织编写，为了方便与技术要求对照，在相关段落前添加表格并罗列相关的技术要求。";
            // add text
            AddParagraphText(filePath, para_1);
            // set paragraph properties
            SetParagraphProperty(filePath);
            // set font properties
            SetParagraphFont(filePath);

            AddParagraphText(filePath, "术语");
            SetParagraphStyle_Title_Level(filePath, "tl3");
            

            string[,] tableData = new string[,]
            {
                { "名称", "描述"},
                { "操作符", "控制系统模型的逻辑处理元素"},
                { "输入流", "控制系统模型的数据输入元素"},
                { "输出流", "控制系统模型的数据输出元素"}
            };
            AddFixedTable(filePath, tableData);

            AddParagraphText(filePath, "基于模型的建模仿真平台功能描述");
            SetParagraphStyle_Title_Level(filePath, "tl2");

            string para_2 = "基于模型的建模仿真平台是面向高安全系统的基于模型的软件设计和开发平台，适合于实时操作系统上应用的开发。它覆盖了原型、设计、调试和测试，能极大地提高软件质量和减少开发时间。";
            AddParagraphText(filePath, para_2);
            SetParagraphProperty(filePath);
            SetParagraphFont(filePath);

            string para_3 = "基于模型的建模仿真平台由基于模型的控制系统软件开发工具、基于模型的人机界面软件开发工具两部分组成。";
            AddParagraphText(filePath, para_3);
            SetParagraphProperty(filePath);
            SetParagraphFont(filePath);

            AddParagraphText(filePath, "基于模型的控制系统软件开发工具");
            SetParagraphStyle_Title_Level(filePath, "tl3");

            string para_4 = "基于模型的控制系统软件开发工具用于控制软件的设计，具有数据流的搭建能力。基于模型的控制系统软件开发工具提供的建模机制都建立在严格的数学模型基础之上，具有严格的数学语义，它们保证了设计模型的精确性、完整性、一致性和无二义性。基于模型的控制系统软件开发工具的模型就是需求的一种明确、无歧义的表达方式。因此，它可以作为一种良好的介质来实现不同项目组、制造商与供应商之间的需求交流。";
            AddParagraphText(filePath, para_4);
            SetParagraphProperty(filePath);
            SetParagraphFont(filePath);

            AddImg(filePath);
        }

        /// <summary>
        /// Create a new docx
        /// </summary>
        /// <param name="filePath"></param>
        public static void CreateDocx(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Add a main document part.
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
            }
        }

        /// <summary>
        /// Insert text into a paragraph
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="text"></param>
        public static void AddParagraphText(string filePath, string text)
        {
            using(WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                // Get docx's MainDocumentPart
                MainDocumentPart mainPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();

                // Get docx's Body
                Body body = mainPart.Document.Body ?? mainPart.Document.AppendChild(new Body());

                // Create a new Paragraph
                Paragraph paragraph = body.AppendChild(new Paragraph());

                // Create a Run
                Run run = paragraph.AppendChild(new Run());

                // Add text
                Text textElement = run.AppendChild(new Text(text));
            }
        }

        /// <summary>
        /// For simply, I just format the paragraph as the first flight indented by 2 characters
        /// and 1.5 times the line spacing.
        /// </summary>
        /// <param name="filePath"></param>
        /// <remarks>
        /// 首行缩进2字符，1.5倍行距
        /// </remarks>
        public static void SetParagraphProperty(string filePath)
        {
            using(WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                try
                {
                    Paragraph lastParagraph = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().Last();

                    ParagraphProperties paragraphProperties = lastParagraph.PrependChild(new ParagraphProperties());

                    // set indent
                    Indentation indentation = paragraphProperties.AppendChild(new Indentation());
                    // 2 characters
                    indentation.FirstLine = "440";
                    indentation.FirstLineChars = 200;

                    // set row spacing
                    SpacingBetweenLines spacingBetweenLines = paragraphProperties.AppendChild(new SpacingBetweenLines());
                    // 1.5 times row spacing
                    spacingBetweenLines.Line = "360";
                    spacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <remarks>
        /// 中文：宋体；英文：Times New Roman， 5号
        /// </remarks>
        public static void SetParagraphFont(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                try
                {
                    Run run = doc.MainDocumentPart.Document.Descendants<Run>().Last();

                    RunProperties runProperties = run.PrependChild(new RunProperties());

                    FontSize fontSize = runProperties.AppendChild(new FontSize());
                    fontSize.Val = "21";

                    RunFonts runFonts = runProperties.AppendChild(new RunFonts());
                    runFonts.Ascii = "Times New Roman";
                    runFonts.EastAsia = "宋体";
                    runFonts.HighAnsi = "Times New Roman";
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// For simply, I add a fixed table.
        /// </summary>
        /// <param name="filePath"></param>
        public static void AddFixedTable(string filePath, string[,] data)
        {
            using (var document = WordprocessingDocument.Open(filePath, true))
            {
                var doc = document.MainDocumentPart.Document;

                Table table = doc.Body.AppendChild(new Table());

                TableProperties props = table.PrependChild(new TableProperties());

                TableBorders tableBorders = props.AppendChild(new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }
                ));

                for (var i = 0; i < data.GetLength(0); i++)
                {
                    var tr = table.AppendChild(new TableRow());

                    for(var j = 0; j < data.GetLength(1); j++)
                    {
                        var tc = tr.AppendChild(new TableCell());
                        tc.Append(new Paragraph(new Run(new Text(data[i, j].ToString()))));

                        Run run = tc.Elements<Paragraph>().First().Elements<Run>().First();

                        RunProperties runProperties = run.PrependChild(new RunProperties());

                        runProperties.AppendChild(new FontSize()
                        {
                            Val = "20"
                        });
                        runProperties.AppendChild(new RunFonts()
                        {
                            EastAsia = "宋体"
                        });

                        // Bold the title
                        if (i == 0)
                        {
                            Bold bold = runProperties.AppendChild(new Bold());
                        }

                        double widthInDxa =  7.23 * 567; // 1cm ≈ 567 Dxa

                        tc.AppendChild(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width= widthInDxa.ToString() }));
                    }
                }
            }
        }

        public static void SetParagraphStyle_Title_Level(string filePath, string styleId)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                Paragraph paragraph = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().Last();

                if(paragraph.Elements<ParagraphProperties>().Count() == 0)
                {
                    paragraph.PrependChild<ParagraphProperties>(new ParagraphProperties());
                }

                ParagraphProperties pPr = paragraph.ParagraphProperties;

                if(pPr.ParagraphStyleId == null)
                {
                    pPr.ParagraphStyleId = new ParagraphStyleId();
                }

                pPr.ParagraphStyleId.Val = styleId;

                doc.MainDocumentPart.Document.Save();
            }
        }

        public static void AddTitleSytle(string filePath)
        {
            using(WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                Styles? styles = doc?.MainDocumentPart?.StyleDefinitionsPart?.Styles ??
                     doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>().Styles;

                if(styles == null)
                {
                    styles = new Styles();
                    doc.MainDocumentPart.StyleDefinitionsPart.Styles = styles;
                }

                // ------ title 1 ------
                Style style_title_1 = styles.AppendChild(new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "tl1",
                    CustomStyle = true
                });

                style_title_1.Append(new Aliases()
                {
                    Val = "自定义标题1"
                });
                style_title_1.Append(new StyleName()
                {
                    Val = "defTitle_1"
                });

                ParagraphProperties pPr_title_1 = style_title_1.AppendChild(new ParagraphProperties());
                pPr_title_1.Append(new Justification()
                {
                    Val = JustificationValues.Center
                });

                pPr_title_1.Append(new OutlineLevel()
                {
                    Val = 1
                });

                pPr_title_1.Append(new SpacingBetweenLines()
                {
                    Line = "400",
                    LineRule = LineSpacingRuleValues.Auto
                });

                StyleRunProperties styleRunProperties_title_1 = style_title_1.AppendChild(new StyleRunProperties());
                styleRunProperties_title_1.Append(new Bold());
                styleRunProperties_title_1.Append(new RunFonts()
                {
                    Ascii = "Times New Roman",
                    EastAsia = "黑体"
                });
                styleRunProperties_title_1.Append(new FontSize()
                {
                    Val = "48"
                });

                // ------ title 2 ------
                Style style_title_2 = styles.AppendChild(new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "tl2",
                    CustomStyle = true
                });

                style_title_2.Append(new Aliases()
                {
                    Val = "自定义标题2"
                });
                style_title_2.Append(new StyleName()
                {
                    Val = "defTitle_2"
                });

                ParagraphProperties pPr_title_2 = style_title_2.AppendChild(new ParagraphProperties());

                pPr_title_2.Append(new OutlineLevel()
                {
                    Val = 2
                });

                pPr_title_2.Append(new SpacingBetweenLines()
                {
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                });

                StyleRunProperties styleRunProperties_title_2 = style_title_2.AppendChild(new StyleRunProperties());
                styleRunProperties_title_2.Append(new RunFonts()
                {
                    Ascii = "Times New Roman",
                    EastAsia = "黑体"
                });
                styleRunProperties_title_2.Append(new FontSize()
                {
                    Val = "32"
                });

                // ------ title 3 ------
                Style style_title_3 = styles.AppendChild(new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "tl3",
                    CustomStyle = true
                });

                style_title_3.Append(new Aliases()
                {
                    Val = "自定义标题3"
                });
                style_title_3.Append(new StyleName()
                {
                    Val = "defTitle_3"
                });

                ParagraphProperties pPr_title_3 = style_title_3.AppendChild(new ParagraphProperties());

                pPr_title_3.Append(new OutlineLevel()
                {
                    Val = 3
                });

                pPr_title_3.Append(new SpacingBetweenLines()
                {
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                });

                StyleRunProperties styleRunProperties_title_3 = style_title_3.AppendChild(new StyleRunProperties());
                styleRunProperties_title_3.Append(new RunFonts()
                {
                    Ascii = "Times New Roman",
                    EastAsia = "黑体"
                });
                styleRunProperties_title_3.Append(new FontSize()
                {
                    Val = "24"
                });
            }
        }

        public static void AddImg(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                ImagePart imgPart = doc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(_imgPath, FileMode.Open))
                {
                    imgPart.FeedData(stream);
                }

                AddImageToBody(doc, doc.MainDocumentPart.GetIdOfPart(imgPart));
            }
        }

        public static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 4939200L, Cy = 2160000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 4939200L, Cy = 2160000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.Document.Body is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
            }

            Paragraph paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center } // Set the alignment to center
                ),
                new Run(element)
            );

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(paragraph);
        }

        public static TValue ValidExpression<TValue>(TValue a, TValue b, string op)
    where TValue : struct, IComparable, IConvertible
        {
            if (op.Equals("*"))
            {
                double result = Convert.ToDouble(a) * Convert.ToDouble(b);

                return (TValue)Convert.ChangeType(result, typeof(TValue));
            }
            else if (op.Equals("/"))
            {
                double result = Convert.ToDouble(a) / Convert.ToDouble(b);

                return (TValue)Convert.ChangeType(result, typeof(TValue));
            }

            throw new InvalidOperationException("Unsupported operation.");
        }
    }
}
