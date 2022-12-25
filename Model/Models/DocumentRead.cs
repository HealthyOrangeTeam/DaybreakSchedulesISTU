using System.Globalization;
using Document = Spire.Doc.Document;
using Section = Spire.Doc.Section;
using Spire.Doc;
using Spire.Doc.Documents;
using Model.Models.Entities;
using Model.Interface;

namespace Model.Models
{
    public class DocumentRead : IDocumentRead
    {
        public IData db { get; set; }


        public string path { get; set; }


        #region нужно!!!!!!!! не удалить
        //public static Image RenderNode(Node node, ImageSaveOptions imageOptions)
        //{
        //    // Run some argument checks.
        //    if (node == null)
        //        throw new ArgumentException("Node cannot be null");

        //    // If no image options are supplied, create default options.

        //    if (imageOptions == null)
        //        imageOptions = new ImageSaveOptions(SaveFormat.Png);

        //    // Store the paper color to be used on the final image and change to transparent.
        //    // This will cause any content around the rendered node to be removed later on.

        //    Color savePaperColor = imageOptions.PaperColor;
        //    imageOptions.PaperColor = Color.Transparent;

        //    // There a bug  which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes,
        //    // find the matching node in the cloned document and render that instead.

        //    Aspose.Words.Document doc = (Aspose.Words.Document)node.Document.Clone(true);
        //    node = doc.GetChild(NodeType.Any, node.Document.GetChildNodes(NodeType.Any, true).IndexOf(node), true);

        //    // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
        //    // the rendered content of the node.

        //    Shape shape = new Shape(doc, ShapeType.TextBox);
        //    Section parentSection = (Section)node.GetAncestor(NodeType.Section);

        //    // Assume that the node cannot be larger than the page in size.
        //    shape.Width = parentSection.PageSetup.PageWidth;
        //    shape.Height = parentSection.PageSetup.PageHeight;

        //    shape.FillColor = Color.Transparent;
        //    // We must make the shape and paper color transparent.
        //    // Don't draw a surronding line on the shape.

        //    shape.Stroked = false;
        //    Node currentNode = node;

        //    // If the node contains block level nodes then just add a copy of these nodes to the shape.
        //    if (currentNode is InlineStory || currentNode is Story)
        //    {
        //        CompositeNode composite = (CompositeNode)currentNode;

        //        foreach (Node childNode in composite.ChildNodes)
        //        {
        //            shape.AppendChild(childNode.Clone(true));
        //        }
        //    }
        //    else
        //    {
        //        // Move up  through the DOM until we find node which is suitable to insert into a Shape(a                node with a parent can contain paragraph, tables the same as a shape).
        //        // Each parent node is cloned on the way up so even a descendant node passed to this            method can be rendered.
        //        // Since we                 are working with the actual nodes of the document we need to clone the target node into the temporary shape.
        //        while (!(currentNode.ParentNode is InlineStory || currentNode.ParentNode is Story ||
        //        currentNode.ParentNode is ShapeBase || currentNode.NodeType == NodeType.Paragraph))
        //        {
        //            CompositeNode parent = (CompositeNode)currentNode.ParentNode.Clone(false);

        //            currentNode = currentNode.ParentNode;
        //            parent.AppendChild(node.Clone(true));
        //            node = parent; // Store this new node to be inserted into the shape.
        //        }

        //        // Add the   node to the shape.
        //        shape.AppendChild(node.Clone(true));
        //    }


        //    // We must add  the shape to the document tree to have it rendered.

        //    parentSection.Body.FirstParagraph.AppendChild(shape);

        //    // Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.

        //    // Retrieve the rendered image and remove the shape from the document.

        //    MemoryStream stream = new MemoryStream();

        //    shape.GetShapeRenderer().Save(stream, imageOptions);

        //    shape.Remove();


        //    Bitmap croppedImage;

        //    // Load the image into a new bitmap.
        //    using (Bitmap renderedImage = new Bitmap(stream))
        //    {
        //        // Extract the actual content of the image by cropping transparent space around
        //        // the rendered shape.
        //        Rectangle cropRectangle = FindBoundingBoxAroundNode(renderedImage);
        //        croppedImage = new Bitmap(cropRectangle.Width, cropRectangle.Height);
        //        croppedImage.SetResolution(imageOptions.HorizontalResolution, imageOptions.VerticalResolution);

        //        // Create the final image with the proper background color.
        //        using (Graphics g = Graphics.FromImage(croppedImage))
        //        {
        //            g.Clear(savePaperColor);
        //            g.DrawImage(renderedImage, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height), cropRectangle.X, cropRectangle.Y, cropRectangle.Width, cropRectangle.Height, GraphicsUnit.Pixel);
        //        }
        //    }

        //    return croppedImage;

        //}
        #endregion



        public void Parse(string path)//TODO: Говно но работает
        {
            this.path = path;

            var GroupItem = new Group();
            GroupItem.Name = Path.GetFileNameWithoutExtension(path);
            GroupItem.PathToDocument = path;

            db.Add(GroupItem);

            Document doc = new Document();
            doc.LoadFromFile(path);
            foreach (Table item in doc.Sections[0].Tables)
            {
                Document doc2 = new Document();
                Section section = doc2.AddSection();

                string end_path = "temp/main";
                string path_week = end_path + "/" + Path.GetFileNameWithoutExtension(path);

                try
                {
                    section.PageSetup.PageSize = new System.Drawing.SizeF(1200 / 72 * 60, 800 / 72 * 60);
                    section.PageSetup.Margins.All = 0f;
                    section.PageSetup.Margins.Left = 51f;
                    section.PageSetup.Orientation = PageOrientation.Landscape;


                    Table copied_Table = item.Clone();
                    doc2.Sections[0].Tables.Add(copied_Table);

                    if (!Directory.Exists(end_path))
                    {
                        Directory.CreateDirectory(end_path);
                        Console.WriteLine($"Create dir {end_path}");
                    }

                    if (!Directory.Exists(path_week))
                    {
                        Directory.CreateDirectory(path_week);
                        Console.WriteLine($"Create dir {path_week}");

                    }


                    if (item.Rows[1].Cells[0].Paragraphs.Count > 1)
                    {

                        var text = item.Rows[1].Cells[0].Paragraphs[1].Text;
                        DateOnly date = DateOnly.Parse(text.Remove(text.Length - 2));
                        var cal = new GregorianCalendar();
                        int week = cal.GetWeekOfYear(date.ToDateTime(TimeOnly.Parse("10:00 PM")), CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

                        var path_docx = $"{end_path}/{Path.GetFileNameWithoutExtension(path)}_result_{week}.docx";

                        var path_jpg = $"{path_week}/_{date.Month}_{week}.jpg";

                        doc2.SaveToFile(path_docx, FileFormat.Docx);
                        var converter = new GroupDocs.Conversion.Converter(path_docx);
                        var convertOptions = converter.GetPossibleConversions()["jpg"].ConvertOptions;

                        converter.Convert(path_jpg, convertOptions);

                        doc2.Close();


                        Console.WriteLine($"Docx to img {path_docx} - {path_jpg}");


                        var weekItem = new Week() { Group = GroupItem, PathToImg = path_jpg, DateTime = date.ToDateTime(TimeOnly.Parse("10:00 PM")) };
                        db.Add(weekItem);
                    }
                }
                catch
                {
                    section.PageSetup.Margins.All = 0f;
                    section.PageSetup.Margins.Left = 51f;
                    section.PageSetup.PageSize = new System.Drawing.SizeF(1200 / 72 * 100, 800 / 72 * 100);
                    section.PageSetup.Orientation = PageOrientation.Portrait;
                    Table copied_Table = item.Clone();
                    doc2.Sections[0].Tables.Add(copied_Table);

                    if (!Directory.Exists(end_path))
                    {
                        Directory.CreateDirectory(end_path);
                        Console.WriteLine($"Create dir {end_path}");
                    }

                    if (!Directory.Exists(path_week))
                    {
                        Directory.CreateDirectory(path_week);
                        Console.WriteLine($"Create dir {path_week}");

                    }

                    if (item.Rows[1].Cells[0].Paragraphs.Count > 1)
                    {

                        var text = item.Rows[1].Cells[0].Paragraphs[1].Text;
                        DateOnly date = DateOnly.Parse(text.Remove(text.Length - 2));
                        var cal = new GregorianCalendar();
                        int week = cal.GetWeekOfYear(date.ToDateTime(TimeOnly.Parse("10:00 PM")), CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

                        var path_docx = $"{end_path}/{Path.GetFileNameWithoutExtension(path)}_result_{week}.docx";
                        var path_jpg = $"{path_week}/_{date.Month}_{week}.jpg";

                        doc2.SaveToFile(path_docx, FileFormat.Docx);
                        var converter = new GroupDocs.Conversion.Converter(path_docx);
                        var convertOptions = converter.GetPossibleConversions()["jpg"].ConvertOptions;
                        converter.Convert(path_jpg, convertOptions);

                        doc2.Close();

                        Console.WriteLine($"Docx to img {path_docx} - {path_jpg}");


                        var weekItem = new Week() { Group = GroupItem, PathToImg = path_jpg, DateTime = date.ToDateTime(TimeOnly.Parse("10:00 PM")) };
                        db.Add(weekItem);
                    }
                }

            }
            doc.Close();
        }

    }
}
