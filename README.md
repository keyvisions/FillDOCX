# FillDOCX

Given a Microsoft DOCX document (“template”) sprinkled with @@\<tag> constructs, and, XML data that includes elements \<tag>, FillDOCX makes a copy of template and replaces, preserving style, the @@\<tag> occurances with the XML element contents. The XML data may be raw, a file or a URL.

The code was written in order to give non programmers an intuitive way of creating automatic fillable templates, they are invited to create a DOCX document sprinkled with @@\<tags> as they best see fit and are asked either to choose or use predefined understandable canonical names in forming the @@\<tag> constructs; these templates are then filled on request in an intranet context with data fetched from a web service.

Note that only the first element with an given tag is used.

`dotnet publish`