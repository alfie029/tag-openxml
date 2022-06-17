// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Newtonsoft.Json;

using var inMemoryOpenXmlDoc = PresentationDocument.Open("Blank.pptx", true);

var presentationPart = inMemoryOpenXmlDoc.PresentationPart ?? inMemoryOpenXmlDoc.AddPresentationPart();
var documentUserTagsPart = presentationPart.UserDefinedTagsPart ?? presentationPart.AddNewPart<UserDefinedTagsPart>();
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
var docTagId = presentationPart.GetIdOfPart(documentUserTagsPart);
presentationPart.Presentation.CustomerDataList ??= new CustomerDataList();
presentationPart.Presentation.CustomerDataList.Append(new CustomerDataTags { Id = docTagId });

presentationPart.SlideParts.ToList().ForEach(slidePart =>
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
});

inMemoryOpenXmlDoc.Save();