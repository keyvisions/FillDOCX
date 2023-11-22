# FillDOCX

Given a Microsoft DOCX document (“template”) sprinkled with @@\<tag> constructs, and, XML data that includes \<tag> elements, FillDOCX makes a copy of template and replaces the @@\<tag> occurances with the corresponding \<tag> XML elements' contents. The XML data may be raw, a file path or a URL.

The code was written in order to give non programmers an intuitive way of creating automatic fillable templates, they are invited to create DOCX documents, designed as they best see fit, sprinkled with @@\<tags> chosen from a given set; these templates are then filled on request in an intranet context with data fetched from a web service.

![From DOCX template to DOCX document](https://github.com/keyvisions/FillDOCX/blob/master/media/visual.jpg "From DOCX template to DOCX document")

## Version 0.9.2
If a @@\<tag> in the template has no associated value an underscore _ is appended to the name of the document indicating that the document is potentially incomplete.

## Version 0.9.1
Revised [hidden] behavior, if it is present in a table row the whole row is eliminated and if it is present in a paragraph the whole paragraph is eliminated.
XML elements that contain HTML are rendered in the document.

## Version 0.9.0
FillDOCX now accepts JSON data, it is transformed in XML and then business as usual (https://github.com/keyvisions/json2xml).

## Version 0.8.2
Fixed error reised when --novalue was set equal to the empty string, i.e., --novalue "": an empty string at the end of the command line is not included in args[].
In Version 0.8.0, when images are replaced, the unconsequential exception "Collection was modified; enumeration operation may not execute." was reised, not anymore. 

## Version 0.8.1
XML Documents are case sensitive, therefore, placeholders should match XML elements. If there is no XML Element associated to a given placeholder, the lowercase XML Element is now searched.
Version 0.7.1 was only applicable to repeating placeholders now also to non-repeating.

## Version 0.8.0
Document images are now replaceable. The DOCX format includes, in the word/media folder, the images embedded in the document, these images are notably named image\<number>.\<extension>. If the XML data includes \<image\<number>> elements their value is interpreted as the path of the new image, image file that replaces the embedded image. Note that FillDOCX does not add images to the DOCX file, it merely replaces existing images.

## Version 0.7.1
Revised [hidden] behaviour, if it is present inside a table the whole table is not rendered.

## Version 0.7.0
Solved bug associated to @@\<tag>.\<tag> construct that referred to data elements with a single child.
Added attribute hidden="true": when an element is assigned this attribute, its content will not be rendered. See also [hidden] Version 0.4.0.

## Version 0.6.0
Opted for FreeSpire.Doc for PDF generation: allows the creation of PDF with up to 3 pages before inserting an "unlicensed" text.

## Version 0.5.0
Added overwrite flag, when set, overwrites the destination file if it exists, by default it will not overwrite.

## Version 0.4.0
Fixed error raised by empty XML elements.
If the text [hidden] appears in a table row with @@\<tag>, the table row is not rendered: practical when a @@\<tag> is set to [hidden].
@@\<tag> should ALWAYS be followed by a non alphanumeric character: DOCX documents inserts spurios markup after the sequence @@.

## Version 0.3.0
Improved the @@\<tag>.\<tag> construct introduced in 0.2.0, now, if placed outside of a table it renders only the first child of @@\<tag>, before, it clobbered the resulting document.docx.

## Version 0.2.0
Introduced the @@\<tag>.\<tag> construct, it _MUST_ be placed inside a table row, with the effect that all children elements of @@\<tag> are rendered in separate rows.
This construct is usefull, for example, to handle purchase orders with multiple lines, see the template.docx template

## Instructions
Clone locally, publish and run

usage: filldocx [\<args_path>] --template \<path> (--xml|--json) (\<path>|\<url>|\<raw>) --destfile \<path> [--pdf] [--overwrite] [--shorttags] [--allowhtml] [--novalue \<string>]

`$ git clone https://github.com/keyvisions/FillDOCX.git`

`$ dotnet restore`

`$ dotnet publish --configuration Release`

`$ dotnet run -t ./template.docx -x ./data.xml -d ./document.docx --pdf`

PDFs can be generated with [Spire.Doc](https://www.e-iceblue.com/Introduce/word-for-net-introduce.html)
