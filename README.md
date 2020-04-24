# FillDOCX

Given a Microsoft DOCX document (“template”) sprinkled with @@\<tag> constructs, and, XML data that includes elements \<tag>, FillDOCX makes a copy of template and replaces, preserving style, the @@\<tag> occurances with the XML element contents. The XML data may be raw, a file or a URL.

The code was written in order to give non programmers an intuitive way of creating automatic fillable templates, they are invited to create a DOCX document sprinkled with @@\<tags> as they best see fit and are asked either to choose or use predefined understandable canonical names in forming the @@\<tag> constructs; these templates are then filled on request in an intranet context with data fetched from a web service.

Note that only the first element with an given tag is used.

## Version 0.3.0
Introduced the @@\<tag>.\<tag> construct, if placed inside a table row then all children elements of @@\<tag> are rendered in separate rows, else, only the first child of @@\<tag> is rendered.

This construct is usefull, for example, to handle purchase orders with multiple lines, see the order.docx template

## Instructions
Clone locally, publish and run

usage: filldocx --template \<path> --xml {\<path>|\<url>|\<raw>} --destfile \<path> [--pdf] [--shorttags] [--novalue \<string>]

`$ git clone https://github.com/keyvisions/FillDOCX.git`

`$ dotnet publish`

`$ dotnet run -t ./order.docx -x ./data.xml -d ./document.docx --pdf`

PDFs can be generated with either [Spire.Doc](https://www.nuget.org/packages/Spire.Doc/) or [sautinsoft.document](https://www.nuget.org/packages/sautinsoft.document/) package.
