// file: ConvertPptx.cs
// last update: 09/11/23

using System;
using System.Text;

using System.Management.Automation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

public enum OutputFormat
{
    Yaml,
    Xml
}

[Cmdlet("ConvertTo", "Text")]
[OutputType(typeof(string))]
public class ConvertPptx2Text : PSCmdlet
{

    private string _PptxPath;

   [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath")]
    [ValidateNotNullOrEmpty()]
 
 
    public string? PptxPath
    {
        get { return _PptxPath; }
        set {
             _PptxPath = value; 
            ProviderInfo provider;
            PSDriveInfo drive;
            var providerPath = this.SessionState.Path.GetUnresolvedProviderPathFromPSPath(_PptxPath, out provider, out drive);
            WriteDebug($"*** {providerPath}");
            _PptxPath = providerPath;
        }
    }

    [Parameter()]
    public OutputFormat OutputFormat { get; set; }

    [Parameter()]
    public SwitchParameter SlidesCount { get; set; }

    [Parameter()]
   public SwitchParameter SlideTitles { get; set; }

    // https://learn.microsoft.com/en-us/office/open-xml/how-to-get-the-titles-of-all-the-slides-in-a-presentation
    
    /// <summary>
    /// Get the number of slides in the presentation
    /// </summary>
    /// <param name="PptxPath"></param>
    /// <returns></returns>
    private int CountSlides(string PptxPath)
    {
        using (PresentationDocument document = PresentationDocument.Open(PptxPath, false))
        {
            return CountSlides(document);
        }
    }

    /// <summary>
    /// Get the number of slides in the document
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    private int CountSlides(PresentationDocument document)
    {
        int slidesCount = 0;
        if (document != null) {
            PresentationPart presentationPart = document.PresentationPart;
            if (presentationPart != null) {
                slidesCount = presentationPart.SlideParts.Count();
            }
        }
        return slidesCount;
    }
    /// <summary>
    /// Get the text of a single slide by its index
    /// </summary>
    /// <param name="PptxPath"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    private SlideContent getSlideContent(string PptxPath, int index)
    {
        // a SlideContent object for the title and the topics
        SlideContent content = new SlideContent();

        using (PresentationDocument document = PresentationDocument.Open(PptxPath, false))
        {
            try {
                PresentationPart presentationPart = document.PresentationPart;
                // get the relationship ID of the first slide
                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[index] as SlideId).RelationshipId;
                // get the slide part from the relationship ID
                SlidePart slide = (SlidePart)presentationPart.GetPartById(relId);
                if (slide != null) {
                    // get the inner text of the slide
                    List<Drawing.Text> textList = slide.Slide.Descendants<Drawing.Text>().ToList();
                    if (textList.Count > 0) {
                        content.Title = textList[0].Text;
                        for(int i=1;i<textList.Count;i++) {
                            // textList contains Drawing.Text objects
                            content.Topics.Add(textList[i].Text);
                        }
                    }              
                }
            } catch(SystemException ex) {
                ErrorRecord errorRecord = new ErrorRecord(ex, "ConvertPptx2Text.GetSlideContent", ErrorCategory.InvalidOperation, null);
                WriteError(errorRecord);

            }
        }
        return content;
    }

    protected override void ProcessRecord()
    {
        List<SlideContent> slideContentList = new List<SlideContent>();
        try
        {
            WriteVerbose($"Start processing {PptxPath}");
            int slideCount = CountSlides(PptxPath);
            if (SlidesCount) {
                this.WriteObject(slideCount);            
            } else {
                for (int i = 0; i < slideCount; i++) {
                    slideContentList.Add(getSlideContent(PptxPath, i)); 
                }
            }
            if (OutputFormat == OutputFormat.Yaml) {
                // Put every object into the pipeline
                var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
                foreach(SlideContent content in slideContentList)
                {
                    var yamlText = serializer.Serialize(content);
                    WriteObject(yamlText);
                }
            } else {
                var ex = new NotSupportedException("Other formats are not done yet, sorry");
                ErrorRecord errorRecord = new ErrorRecord(ex, "PptxGeneral", ErrorCategory.NotImplemented, null);
                WriteError(errorRecord);
            }

       } catch (SystemException ex) {
            ErrorRecord errorRecord = new ErrorRecord(ex, "ConvertPptx2Text.ProcessRecord", ErrorCategory.InvalidOperation, null);
            this.WriteError(errorRecord);
        }
    }
}