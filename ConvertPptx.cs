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
            // Check if the file exists
            if (!File.Exists(_PptxPath))
            {
                System.IO.FileNotFoundException ex = new FileNotFoundException($"{providerPath} does not exist or has never existed or whatever.");
                ErrorRecord errorRecord = new ErrorRecord(ex, "ConvertPptx2Text.PptxPath", ErrorCategory.InvalidOperation, null);
                WriteError(errorRecord);
            }
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
    /// Determines whether the shape is a title shape
    /// </summary>
    /// <param name="shape"></param>
    /// <returns></returns>
    private static bool IsTitleShape(Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
        if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
        {
            switch ((PlaceholderValues)placeholderShape.Type)
            {
                // Any title shape
                case PlaceholderValues.Title:
                // Centered title
                case PlaceholderValues.CenteredTitle:
                    return true;
                default:
                    return false;
            }
        }
        return false;
    }

    /// <summary>
    /// gets the text of a Shape
    /// </summary>
    /// <param name="shape"></param>
    /// <returns></returns>
    private string getShapeText(Shape shape) 
    {
        StringBuilder paragraphText = new StringBuilder();
        foreach (var paragraph in shape.TextBody.Descendants<Drawing.Paragraph>())
        {
            foreach (var text in paragraph.Descendants<Drawing.Text>())
            {
                paragraphText.Append(text.Text);
            }
            paragraphText.Append("\n");
        }
        return paragraphText.ToString();
    }

    /// <summary>
    /// Returns the title of a slide or an empty string
    /// </summary>
    /// <param name="slidePart"></param>
    /// <returns></returns>
    private String getSlideTitle(SlidePart slidePart)
    {
        try {
            // Declare a paragraph separator
            string paragraphSeparator = null;
            if (slidePart.Slide != null) {
                // Find all the title shapes
                var shapes = from shape in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                            where IsTitleShape(shape)
                            select shape;

                StringBuilder paragraphText = new StringBuilder();

                foreach (var shape in shapes)
                {
                    // Get the text in each paragraph in this shape
                    foreach (var paragraph in shape.TextBody.Descendants<Drawing.Paragraph>())
                    {
                        // Add a line break
                        paragraphText.Append(paragraphSeparator);
                        foreach (var text in paragraph.Descendants<Drawing.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }
                        paragraphSeparator = "\n";
                    }
                }
                return paragraphText.ToString();
            }
            return string.Empty;
        } catch(SystemException ex) {
            ErrorRecord errorRecord = new ErrorRecord(ex, "ConvertPptx2Text.getSlideTitle", ErrorCategory.InvalidOperation, null);
            WriteError(errorRecord);
            return string.Empty;
        }

    }
    private List<string> getSlideTitles(string PptxPath)
    {
        List<string> titlesList = new List<string>();
        try {
            using (PresentationDocument document = PresentationDocument.Open(PptxPath, false))
            {
                PresentationPart presentationPart = document.PresentationPart;
                if (presentationPart != null && presentationPart.Presentation != null) {
                    Presentation presentation = presentationPart.Presentation;
                    if (presentation.SlideIdList != null) {
                        // Get the title of each slide in the slide order.
                        foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                        {
                            SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                            // Get the slide title
                            string title = getSlideTitle(slidePart);
                            // An empty title can also be added
                            titlesList.Add(title);
                        }
                    }
                }
            }
        } catch(SystemException ex) {
            ErrorRecord errorRecord = new ErrorRecord(ex, "ConvertPptx2Text.getTitles", ErrorCategory.InvalidOperation, null);
            WriteError(errorRecord);
        }
        return titlesList;
    }

    /// <summary>
    /// Get the number of slides in the presentation
    /// </summary>
    /// <param name="PptxPath"></param>
    /// <returns></returns>
    private int countSlides(string PptxPath)
    {
        using (PresentationDocument document = PresentationDocument.Open(PptxPath, false))
        {
            return countSlides(document);
        }
    }

    /// <summary>
    /// Get the number of slides in the document
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    private int countSlides(PresentationDocument document)
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
            string paragraphSeparator = "\n";
            try {
                PresentationPart presentationPart = document.PresentationPart;
                // get the relationship ID of the first slide
                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[index] as SlideId).RelationshipId;
                // get the slide part from the relationship ID
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relId);
                if (slidePart != null) {
                    foreach (Shape shape in slidePart.Slide.Descendants<Shape>()) {
                        if (shape.TextBody != null) {
                            // is it the title shape?
                            if (IsTitleShape(shape)) {
                                string shapeText = "";
                                foreach (Drawing.Paragraph paragraph in shape.TextBody.Descendants<Drawing.Paragraph>())
                                {
                                    shapeText += paragraph.InnerText;
                                }
                                content.Title = shapeText;
                            } else {
                                content.Topics.Add(getShapeText(shape));
                            }
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

    /// <summary>
    /// Processes the pipeline object by object
    /// </summary>
    protected override void ProcessRecord()
    {
        List<SlideContent> slideContentList = new List<SlideContent>();
        try
        {
            WriteVerbose($"Start processing {PptxPath}");
            int slideCount = countSlides(PptxPath);
            // Parameter SlidesCount?
            if (this.MyInvocation.BoundParameters.ContainsKey("SlidesCount")) {
                WriteObject(slideCount);            
                WriteVerbose($"{slideCount} slides collected.");
            // Parameter SlideTitles?
            } else if (this.MyInvocation.BoundParameters.ContainsKey("SlideTitles")) {
                var slideTitles = getSlideTitles(PptxPath);
                WriteObject(slideTitles);
                WriteVerbose($"{slideTitles.Count} slide titles collected.");
            // No switch parameter
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