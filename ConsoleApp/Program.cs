using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Model.Models;
using Model.Models.Data;

MainClass d = new MainClass()
{
    Read = new DocumentRead() {
        db = new Data()
    },
};

d.Read.Parse("temp/test.docx");


#region МОЖЕ НАДА

//Document doc = new Document("temp/test.docx");

//NodeCollection paragraphs = doc.GetChildNodes(NodeType.Table, true);


//int n = 0;
//foreach (var item in paragraphs)
//{
//    Console.WriteLine(paragraphs.Count());
//    Document dstDoc = new Document();
//    //Node table = (Table)item.GetAncestor(NodeType.Table);

//    if (item != null)
//    {
//        NodeImporter importer = new NodeImporter(doc, dstDoc, ImportFormatMode.UseDestinationStyles);
//        Node newnode = importer.ImportNode(item, true);
//        dstDoc.FirstSection.Body.AppendChild(newnode);
//        dstDoc.Save($"test{n}.docx"); 
//        n++;

//        Console.WriteLine($"test{n}.docx");
//    }

//}


//Document doc = new Document(@"temp/КПІ-91.docx");


//List<Image> images = new List<Image>();



//NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
//foreach (Table table in tables)
//{
//    images.Add(RenderNode(table, new ImageSaveOptions(SaveFormat.Png)));
//}


//n = 0;
//foreach (var item in images)
//{
//    item.Save($"temp/main/test{n}.png", ImageFormat.Png);
//    n++;
//}









static Image RenderNode(Node node, ImageSaveOptions imageOptions)
{
    // Run some argument checks.
    if (node == null)
        throw new ArgumentException("Node cannot be null");

    // If no image options are supplied, create default options.

    if (imageOptions == null)
        imageOptions = new ImageSaveOptions(SaveFormat.Png);

    // Store the paper color to be used on the final image and change to transparent.
    // This will cause any content around the rendered node to be removed later on.

    Color savePaperColor = imageOptions.PaperColor;
    imageOptions.PaperColor = Color.Transparent;

    // There a bug  which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes,
    // find the matching node in the cloned document and render that instead.

    Document doc = (Document)node.Document.Clone(true);
    node = doc.GetChild(NodeType.Any, node.Document.GetChildNodes(NodeType.Any, true).IndexOf(node), true);

    // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
    // the rendered content of the node.

    Shape shape = new Shape(doc, ShapeType.TextBox);
    Section parentSection = (Section)node.GetAncestor(NodeType.Section);

    // Assume that the node cannot be larger than the page in size.
    shape.Width = parentSection.PageSetup.PageWidth ;
    shape.Height = parentSection.PageSetup.PageHeight ;

    shape.FillColor = Color.Transparent;
    // We must make the shape and paper color transparent.
    // Don't draw a surronding line on the shape.

    shape.Stroked = false;
    Node currentNode = node;

    // If the node contains block level nodes then just add a copy of these nodes to the shape.
    if (currentNode is InlineStory || currentNode is Story)
    {
        CompositeNode composite = (CompositeNode)currentNode;

        foreach (Node childNode in composite.ChildNodes)
        {
            shape.AppendChild(childNode.Clone(true));
        }
    }
    else
    {
        // Move up  through the DOM until we find node which is suitable to insert into a Shape(a                node with a parent can contain paragraph, tables the same as a shape).
        // Each parent node is cloned on the way up so even a descendant node passed to this            method can be rendered.
        // Since we                 are working with the actual nodes of the document we need to clone the target node into the temporary shape.
        while (!(currentNode.ParentNode is InlineStory || currentNode.ParentNode is Story ||
        currentNode.ParentNode is ShapeBase || currentNode.NodeType == NodeType.Paragraph))
        {
            CompositeNode parent = (CompositeNode)currentNode.ParentNode.Clone(false);

            currentNode = currentNode.ParentNode;
            parent.AppendChild(node.Clone(true));
            node = parent; // Store this new node to be inserted into the shape.
        }

        // Add the   node to the shape.
        shape.AppendChild(node.Clone(true));
    }


    // We must add  the shape to the document tree to have it rendered.

    parentSection.Body.FirstParagraph.AppendChild(shape);

    // Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.

    // Retrieve the rendered image and remove the shape from the document.

    MemoryStream stream = new MemoryStream();

    shape.GetShapeRenderer().Save(stream, imageOptions);

    shape.Remove();


    Bitmap croppedImage;

    // Load the image into a new bitmap.
    using (Bitmap renderedImage = new Bitmap(stream))
    {
        // Extract the actual content of the image by cropping transparent space around
        // the rendered shape.
        Rectangle cropRectangle = FindBoundingBoxAroundNode(renderedImage);
        croppedImage = new Bitmap(cropRectangle.Width, cropRectangle.Height);
        croppedImage.SetResolution(imageOptions.HorizontalResolution, imageOptions.VerticalResolution);

        // Create the final image with the proper background color.
        using (Graphics g = Graphics.FromImage(croppedImage))
        {
            g.Clear(savePaperColor);
            g.DrawImage(renderedImage, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height), cropRectangle.X, cropRectangle.Y, cropRectangle.Width, cropRectangle.Height, GraphicsUnit.Pixel);
        }
    }

    return croppedImage;

}

///
/// Finds the minimum bounding box around non-transparent pixels in a  Bitmap.
///
static Rectangle FindBoundingBoxAroundNode(Bitmap originalBitmap)
{
    Point min = new Point(int.MaxValue, int.MaxValue);
    Point max = new Point(int.MinValue, int.MinValue);

    for (int x = 0; x < originalBitmap.Width; ++x)
    {
        for (int y = 0; y < originalBitmap.Height; ++y)
        {
            // Note  that you can speed up this part of the algorithm by using LockBits and unsafe code instead of GetPixel.
            Color pixelColor = originalBitmap.GetPixel(x, y);

            // For each pixel that is not transparent calculate the bounding box around it.
            if (pixelColor.ToArgb() != Color.Empty.ToArgb())
            {
                min.X = Math.Min(x, min.X);
                min.Y = Math.Min(y, min.Y);
                max.X = Math.Max(x, max.X);
                max.Y = Math.Max(y, max.Y);
            }
        }
    }

    // Add one pixel to the width and height to avoid clipping.
    return new Rectangle(min.X , min.Y, (max.X - min.X) + 1, (max.Y - min.Y) + 1);
}


#endregion