# C# library to generate PowerPoint files from templates with basic HTML parsing functionality

This library uses the Office Open XML format (pptx) through the [Open XML SDK 2.0 for Microsoft Office](http://www.microsoft.com/en-us/download/details.aspx?id=5124).
Fork extends dependency list by [MariGold.HtmlParser](https://github.com/kannan-ar/MariGold.HtmlParser) and [MariGold.OpenXHTML](https://github.com/kannan-ar/MariGold.OpenXHTML).
Generated files should be opened using Microsoft PowerPoint >= 2010.

PptxTemplater handles:
- Text tags
- Slides (add/remove)
- Slide notes
- Tables (add/remove columns)
- Pictures
- Hyperlinks (this fork feature)
- basic HTML parsing (this fork feature)

## Example

Create a PowerPoint template with two slides and inserts tags (`{{hello}}`, `{{bonjour}}`, `{{hola}}`) in it,
then generate the final PowerPoint file using the following code:

```C#
const string srcFileName = "template.pptx";
const string dstFileName = "final.pptx";
File.Delete(dstFileName);
File.Copy(srcFileName, dstFileName);

Pptx pptx = new Pptx(dstFileName, FileAccess.ReadWrite);
int nbSlides = pptx.SlidesCount();
Assert.AreEqual(2, nbSlides);

// First slide
{
    PptxSlide slide = pptx.GetSlide(0);
    slide.ReplaceTag("{{hello}}", "HELLO HOW ARE YOU?", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{bonjour}}", "BONJOUR TOUT LE MONDE", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{hola}}", "HOLA MAMA QUE TAL?", PptxSlide.ReplacementType.Global);
}

// Second slide
{
    PptxSlide slide = pptx.GetSlide(1);
    slide.ReplaceTag("{{hello}}", "H", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{bonjour}}", "B", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{hola}}", "H", PptxSlide.ReplacementType.Global);
}
```

## Implementation

The source code is clean, documented, tested and should be stable.
A good amount of unit tests come with the source code.

## Fork remarks

First of all my appreciation to original author's effort - great work! Still (after years) very handful piece of code.

Extensions to original code provides basic HTML parsing and hyperlinks. I am not planning to provide further development (at least not on regular basis) except Known Bugs (see below).
I have added parts required for my other project and just sharing with community.
However if you see room for improvements/fixes - pull requests please (original repo is closed and archived). 
**Example task for contributors** - provide parsing <img/> HTML tag implementation.
Enclosed unit tests are the best way to start even for those who are not familiar with OpenXML structure - you should be able to get a grip on that very quickly. 

Links for those of you who are eager to go deeper with the topic:
* [How do I...](https://docs.microsoft.com/en-us/office/open-xml/how-do-i)
* [Structure of a PresentationML document (Open XML SDK)](https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document)
* [Fundamentals and Markup Language Reference](https://www.iso.org/standard/71691.html)

Last position is just to give you a clue how broad is that topic. This repository just scratches the surface but even though it provides quite useful set of functionallity.

## Known bugs

none
