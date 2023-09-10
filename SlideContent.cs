// file: SlideContent.cs

using System.Collections.Generic;

public class SlideContent
{
    public string? Title {get; set;}
    public List<string>? Topics {get;set;}

    public SlideContent()
    {
        this.Topics = new List<string>();
    }
}