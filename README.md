# PoshPptx2Text
A PowerShell module for extracting all text from a Pptx file

This is a simple PowerShell cmdlet that I use for my advance PowerShell training classes.

It extracts all the text from a PowerPoint presentation file.

Since it uses the *DocumentFormat.dll* assembly (*OpenXml*) PowerPoint as an application is **not** involed and does **not** have to be installed.

The Cmdlet will work on Linux and MacOs too (but not tested yet).

The output is either *Yaml* or *XML* (not implemented yet).

You find the explanation for working with the presentation types from the *OpenXML SDK* mostly in the Microsoft documentation:

[https://learn.microsoft.com/en-us/office/open-xml/presentations](https://learn.microsoft.com/en-us/office/open-xml/presentations)

But I also have spend some thoughts on my own;)

By the way, ChatGPT did **not** write the code. I tried to use if out of curiosity of course, but the generated code did not work due to the minor fact that the code enumerated the *TextElement* descendants and not the *Paragraph* descendants.

I could have pressed ChatGPT more for a correct solution, but it was not worth the effort.

The current solution is not perfect either, but its supposed to be more a demonstration for a cmdlet project than a solid cmdlet. And there are probably some modules on the PowerShell gallery that really do the job.

The cmdlet will work with PowerShell 7.x but **not** with *Windows PowerShell*. It would be possible of course by compiling the cs file with the C# compiler from the .Net Framework.

At the moment, the dlls are not copied automatically after a *dotnet publish* from the publish subdirectory into the module directory. This has to be done by copying the dll files. The module directory also needs the *DocumentFormat.OpenXml.dll* file. So, the deployment is not very agile at the moment. But again, this is for education purpose only.

Peter Monadjemi
