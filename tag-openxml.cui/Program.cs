// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Newtonsoft.Json;
using Extension = DocumentFormat.OpenXml.Drawing.Extension;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using NonVisualPictureProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties;
using NonVisualShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties;
using Picture = DocumentFormat.OpenXml.Presentation.Picture;
using Position = DocumentFormat.OpenXml.Presentation.Position;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using Text = DocumentFormat.OpenXml.Presentation.Text;

const string SeismicLivedocExtensionUri = "{721CA3E0-AD4C-468F-9A70-C8A3B087B49E}";

using var inMemoryOpenXmlDoc = PresentationDocument.Open("Blank.pptx", true);

var presentationPart = inMemoryOpenXmlDoc.PresentationPart ?? inMemoryOpenXmlDoc.AddPresentationPart();

WriteDocumentTag(presentationPart);

presentationPart.SlideParts.ToList().ForEach(slidePart =>
{
    WriteSlideTag(slidePart);
    WriteIdentityForVariable(slidePart);
});

inMemoryOpenXmlDoc.Save();


void WriteDocumentTag(PresentationPart presentationPart1)
{
    var documentUserTagsPart =
        presentationPart1.UserDefinedTagsPart ?? presentationPart1.AddNewPart<UserDefinedTagsPart>();
    documentUserTagsPart.TagList ??= new TagList();

    documentUserTagsPart.TagList.Append(new Tag()
    {
        Name = "ContosoInstance",
        Val = JsonConvert.SerializeObject(new
        {
            InstanceId = Guid.Parse("fabf8c6b-5bc8-49bc-bdaf-dfed5ac295ff"),
            Generator = "LiveDoc",
            GenerationId = Guid.Parse("3c275dde-7f4f-49d2-aa85-5fd731ae4b08"),
            Storages = Array.Empty<object>()
        }),
    });

    // Save those tag relationship to presentation customerDataList
    var docTagId = presentationPart1.GetIdOfPart(documentUserTagsPart);
    presentationPart1.Presentation.CustomerDataList ??= new CustomerDataList();
    presentationPart1.Presentation.CustomerDataList.Append(new CustomerDataTags { Id = docTagId });
}

void WriteSlideTag(SlidePart slidePart)
{
    var slideUserTagsPart =
        slidePart.UserDefinedTagsParts.FirstOrDefault() ?? slidePart.AddNewPart<UserDefinedTagsPart>();
    slideUserTagsPart.TagList ??= new TagList();
    slideUserTagsPart.TagList.Append(new Tag()
    {
        Name = "ContosoPage",
        Val = JsonConvert.SerializeObject(new
        {
            PageId = Guid.Parse("10d8a590-e774-4169-abcb-9cd8eb50c227"),
            Index = slideUserTagsPart.TagList.Count(),
            Reference = new
            {
                Repository = "Workspace",
                ContentId = Guid.Parse("c9a99445-567d-4fd7-9bc3-237580ff5574"),
                ContentVersionId = Guid.Parse("4dcd572f-ffca-4f96-a84a-3af941e0b798"),
                ContentVersion = "2.0",
                Page = new
                {
                    PageId = Guid.Parse("ec249b82-4454-44fe-89f0-856d46c91d2e"),
                    Index = 3,
                    slideId = 25
                },
            }
        }),
    });

    // Save tag to slide and correct the relationship to CustomerDataList 
    var tagId = slidePart.GetIdOfPart(slideUserTagsPart);
    slidePart.Slide.CommonSlideData ??= new CommonSlideData();
    slidePart.Slide.CommonSlideData.CustomerDataList ??= new CustomerDataList();
    slidePart.Slide.CommonSlideData.CustomerDataList.Append(new CustomerDataTags { Id = tagId });

    slidePart.Slide.Save();
}

void WriteIdentityForVariable(SlidePart slidePart)
{
    slidePart.Slide.CommonSlideData ??= new CommonSlideData();
    slidePart.Slide.CommonSlideData.ShapeTree?.Where(s => s is Shape)
        .Cast<Shape>()
        .ToList()
        .ForEach(shape =>
        {
            var nvSpPr = shape.NonVisualShapeProperties ??= new NonVisualShapeProperties();
            var cNvPr = nvSpPr.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
            var extLst = cNvPr.NonVisualDrawingPropertiesExtensionList ??=
                new NonVisualDrawingPropertiesExtensionList();

            extLst.Append(new Extension(
                new Comment()
                {
                    Position = new Position() { X = 50, Y = 50 },
                    Text = new Text($"text variable info: {shape.InnerText}")
                }
            ) { Uri = SeismicLivedocExtensionUri });
        });

    slidePart.Slide.CommonSlideData.ShapeTree?.Where(pic => pic is Picture)
        .Cast<Picture>()
        .ToList()
        .ForEach(pic =>
        {
            var nvPicPr = pic.NonVisualPictureProperties ??= new NonVisualPictureProperties();
            var cNvPr = nvPicPr.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
            var extLst = cNvPr.NonVisualDrawingPropertiesExtensionList ??=
                new NonVisualDrawingPropertiesExtensionList();

            extLst.Append(new Extension(
                new Comment()
                {
                    Position = new Position() { X = 50, Y = 50 },
                    Text = new Text($"picture variable info: {cNvPr.Name}")
                }
            ) { Uri = SeismicLivedocExtensionUri });
        });

    slidePart.Slide.Save();
}