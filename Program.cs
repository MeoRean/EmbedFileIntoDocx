using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using OpenMcdf;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml;
using System.IO.Packaging;

namespace EmbedFileIntoDocx;

internal class Program
{
    static void Main(string[] args)
    {
        WordprocessingDocument doc = WordprocessingDocument.Create("Test.docx", WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body();
        mainPart.Document.Append(body);

        //创建Ole文件（oleObject1.bin)
        CompoundFile cf = new CompoundFile();
        cf.RootStorage.CLSID = new Guid("0003000C-0000-0000-C000-000000000046");
        AddCompObjStream(cf.RootStorage);
        AddOle10NativeStream(cf.RootStorage, "attachment.zip");
        AddObjInfoStream(cf.RootStorage);

        //把Ole文件写入到文档中
        var embeddedPart = mainPart.AddNewPart<EmbeddedObjectPart>("application/vnd.openxmlformats-officedocument.oleObject");
        using (MemoryStream ms = new MemoryStream())
        {
            cf.Save(ms);
            ms.Position = 0;
            embeddedPart.FeedData(ms);
        }
        cf.Close();
        //获取写入Ole文件的ID
        string embeddedObjectPartId = mainPart.GetIdOfPart(embeddedPart);

        //读取图标文件，并写入到文件中
        byte[] iconBytes = File.ReadAllBytes("attachment.emf");
        var imagePart = mainPart.AddImagePart(ImagePartType.Emf);//注意设置对应的格式
        using (BinaryWriter writer = new BinaryWriter(imagePart.GetStream()))
        {
            writer.Write(iconBytes);
        }
        //获取写入图标文件的ID
        string imagePartId = mainPart.GetIdOfPart(imagePart);
        var shapeId = "_" + Guid.NewGuid().ToString("N");
        var objectId = "_" + Guid.NewGuid().ToString("N");
        //把图标文件创建为Shape
        var shape = new Shape(
            new ImageData()
            {
                RelationshipId = imagePartId,
                Title = ""
            })
        {
            Id = shapeId,
            Style = "width:50pt;height:50pt",
            Type = "#_x0000_t75"
        };
        //把ole文件创建为ole对象，ole对象关联了图标的Shape
        var oleObject = new OleObject()
        {
            Type = OleValues.Embed,
            ProgId = "Package",
            Id = embeddedObjectPartId,
            ShapeId = shapeId,
            ObjectId = objectId,
            DrawAspect = OleDrawAspectValues.Content
        };
        //创建一个新的Run，里面包含了一个EmbeddedObject，其包含了shape和ole对象
        var run = new Run(
            new RunProperties(),
            new EmbeddedObject(
                shape,
                oleObject
            )
        );
        //插入到文件中
        var paragraph = new Paragraph(run);

        body.Append(paragraph);
        doc.Save();

    }

    private static void AddOle10NativeStream(CFStorage storage, string filePath)
    {
        short flags1 = 2; 
        short flags2 = 0; 
        short unknown1 = 3; 
        short flags3 = 0; 
        var fileName = System.IO.Path.GetFileName(filePath);
        byte[] fileBytes = File.ReadAllBytes(filePath);
        CFStream ole10NativeStream = storage.AddStream("\x01Ole10Native");
        using MemoryStream oleStream = new MemoryStream();
        using LittleEndianOutputStream leosOut = new LittleEndianOutputStream(oleStream);
        using MemoryStream bos = new MemoryStream();
        using LittleEndianOutputStream leos = new LittleEndianOutputStream(bos);

        leos.WriteShort(flags1);
        leos.Write(Encoding.ASCII.GetBytes(fileName));
        leos.WriteByte(0);
        leos.Write(Encoding.ASCII.GetBytes(fileName));
        leos.WriteByte(0);
        leos.WriteShort(flags2);
        leos.WriteShort(unknown1);
        leos.WriteInt(fileName.Length + 1);
        leos.Write(Encoding.ASCII.GetBytes(fileName));
        leos.WriteByte(0);
        leos.WriteInt(fileBytes.Length);
        leos.Write(fileBytes);
        leos.WriteShort(flags3);

        leosOut.WriteInt((int)bos.Length); // total size
        bos.WriteTo(oleStream);


        byte[] oleBytes = oleStream.ToArray();
        ole10NativeStream.SetData(oleBytes);
    }

    private static void AddCompObjStream(CFStorage storage)
    {
        CFStream compObjStream = storage.AddStream("\u0001CompObj");
        using (MemoryStream ms = new MemoryStream())
        using (BinaryWriter writer = new BinaryWriter(ms))
        {
            writer.Write(new byte[]
            {
                0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0x0C, 0x00, 0x03, 0x00, 0x00,
                0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46, 0x0C, 0x00, 0x00, 0x00, 0x4F, 0x4C,
                0x45, 0x20, 0x50, 0x61, 0x63, 0x6B, 0x61, 0x67, 0x65, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00,
                0x00, 0x50, 0x61, 0x63, 0x6B, 0x61, 0x67, 0x65, 0x00, 0xF4, 0x39, 0xB2, 0x71, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            });
            compObjStream.SetData(ms.ToArray());
        }
    }

    private static void AddObjInfoStream(CFStorage storage)
    {
        CFStream objInfoStream = storage.AddStream("\u0003ObjInfo");
        using (MemoryStream ms = new MemoryStream())
        using (BinaryWriter writer = new BinaryWriter(ms))
        {
            writer.Write(new byte[] { 0x00, 0x00, 0x03, 0x00, 0x01, 0x00 });
            objInfoStream.SetData(ms.ToArray());
        }
    }
}