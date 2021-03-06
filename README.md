# FillDOCX

Given a Microsoft DOCX document (“template”) sprinkled with @@\<tag> constructs, and, XML data that includes elements \<tag>, FillDOCX makes a copy of template and replaces, preserving style, the @@\<tag> occurances with the XML element contents&mdash;add a non alphanumeric character at the end of the name. The XML data may be raw, a file or a URL.

The code was written in order to give non programmers an intuitive way of creating automatic fillable templates, they are invited to create a DOCX document sprinkled with @@\<tags> as they best see fit and are asked either to choose or use predefined understandable canonical names in forming the @@\<tag> constructs; these templates are then filled on request in an intranet context with data fetched from a web service.

Note that only the first element with an given tag is used.

## Version 0.7.1
Revised [hidden] behaviour, if it is present inside a table the whole table is not  rendered.

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
This construct is usefull, for example, to handle purchase orders with multiple lines, see the order.docx template

## Instructions
Clone locally, publish and run

usage: filldocx --template \<path> --xml {\<path>|\<url>|\<raw>} --destfile \<path> [--pdf] [--overwrite] [--shorttags] [--novalue \<string>]

`$ git clone https://github.com/keyvisions/FillDOCX.git`

`$ dotnet add package FreeSpire.Doc`

`$ dotnet publish`

`$ dotnet run -t ./order.docx -x ./data.xml -d ./document.docx --pdf`

PDFs can be generated with the [Spire.Doc](https://www.nuget.org/packages/Spire.Doc/) package.
