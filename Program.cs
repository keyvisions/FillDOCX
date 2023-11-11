using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO.Compression;
using Spire.Doc; // https://www.e-iceblue.com/Introduce/spire-office-for-net-free.html
using System.Security;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Web;
using System.Text;
using DocumentFormat.OpenXml;

namespace FillDOCX
{
    class Program
    {
        private static readonly Regex PLACEHOLDER = new Regex(@"@@([a-z]\w*)\.?([a-z]\w*)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static string Fill(string template, XmlElement data, string novalue = "***", int level = 1)
        {
            if (data.Attributes.GetNamedItem("hidden") != null)
                return "";

            List<string> tags = new List<string>();
            foreach (Match match in PLACEHOLDER.Matches(template))
            {
                if ((level == 1 || data.Name == match.Groups[1].Value) && !tags.Contains(match.Groups[level].Value))
                    tags.Add(match.Groups[level].Value);
            }
            tags.Sort((x, y) => y.Length.CompareTo(x.Length));

            tags.ForEach(tag =>
            {
                XmlNodeList nodes = data.GetElementsByTagName(tag);
                if (nodes.Count == 0)
                    nodes = data.GetElementsByTagName(tag.ToLower());
                string subtemplate = level == 1 ? $"@@{tag}" : $"@@{data.Name}.{tag}", value = novalue;

                if (nodes.Count > 0)
                {
                    if (nodes[0].HasChildNodes && nodes[0].FirstChild.GetType() != typeof(System.Xml.XmlText) && level == 1)
                    {
                        subtemplate = new Regex(@"<w:tr (?:(?!<w:tr ).)*?@@" + Regex.Escape(tag) + @".*?<\/w:tr>", RegexOptions.Compiled).Match(template).Value;
                        if (subtemplate.IndexOf("</w:tr>") < subtemplate.Length - 7)
                        {
                            //                            subtemplate = new Regex(@"<w:t>@@" + Regex.Escape(tag) + @".*?<\/w:t>", RegexOptions.Compiled).Match(template).Value;
                            subtemplate = new Regex(@"<w:t>@@" + tag + @".*?<\/w:t>", RegexOptions.Compiled).Match(template).Value;
                            if (subtemplate == "")
                                return;
                            value += Fill(subtemplate, (XmlElement)nodes[0], novalue, level + 1);
                        }
                        else
                        {
                            foreach (XmlElement node in nodes)
                                value += Fill(subtemplate, node, novalue, level + 1);
                        }
                    }
                    else if (nodes[0].Attributes.GetNamedItem("hidden") != null)
                        value = "";
                    else
                        value = nodes[0].InnerXml;
                }
                if (Regex.Match(tag, @"^image\d+").Success)
                    value = "";

                if (subtemplate != "")
                {
                    if (value.IndexOf("altChunk") != -1)
                        value = HttpUtility.HtmlDecode(value);
                    template = template.Replace(subtemplate, value);
                }
            });
            return template;
        }
        private static string Cleanup(string body)
        {
            Regex PLACEHOLDER = new Regex(@"@@\w+(\.\w+)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            Regex USELESS = new Regex(@"</w:t></w:r><[\s\S]*?(<w:t>|<w:t [\s\S]*?>)(?<whitespace>.{1})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(body);

            int s = -2, i = 0; // i to exit potential infinite loop
            MatchCollection matches = PLACEHOLDER.Matches(xmlDoc.InnerText);
            foreach (Match match in matches.Cast<Match>())
            {
                string tag = match.Value;

                s = body.IndexOf("@@", s + 2);
                while (!body[s..].StartsWith(tag) && i < 10)
                {
                    Match useless = USELESS.Match(body, s);
                    if (!useless.Success || useless.Value[13] == '/' || char.IsWhiteSpace(useless.Groups["whitespace"].Value, 0))
                        break;
                    body = body.Remove(useless.Index, useless.Length - 1);
                    ++i;
                }
            }
            xmlDoc = null;

            return body;
        }
        private static string FillDOCX(string template, string mime, string txt, string destfile, string novalue, bool overwrite = false, bool pdf = false, bool shortTags = false, bool allowHTML = false)
        {
            XmlDocument data = new XmlDocument();
            data.PreserveWhitespace = true;

            switch (mime)
            {
                case "application/xml":
                    try { data.Load(txt); } catch { data.LoadXml(txt); }
                    break;

                case "application/json":
                    try
                    {
                        if (txt.IndexOfAny(Path.GetInvalidPathChars()) == -1)
                        {
                            if (txt.StartsWith("http"))
                            {
                                WebClient client = new WebClient();
                                txt = client.DownloadString(txt);
                                client.Dispose();
                            }
                            else
                            {
                                StreamReader sr = new StreamReader(txt);
                                txt = sr.ReadToEnd();
                                sr.Dispose();
                            }
                        }
#if DEBUG
                        Console.Write(Json2Xml(txt));
#endif
                        data.LoadXml(Json2Xml(txt));
                    }
                    catch (SystemException e)
                    {
                        return String.Format("{0}: {1}", e.GetType().Name, e.Message);
                    }
                    break;

                default:
                    try
                    {
                        throw new ArgumentException("Specify --xml or --json");
                    }
                    catch (SystemException e)
                    {
                        return String.Format("{0}: {1}", e.GetType().Name, e.Message);
                    }
            }

            try
            {
                if (template == destfile)
                    throw new ArgumentException("Template cannot be also the destination");

                if (File.Exists(destfile) && !overwrite)
                    goto pdf; // return destfile;

                if (Path.GetDirectoryName(destfile) != "")
                    Directory.CreateDirectory(Path.GetDirectoryName(destfile)); // Create directory if it does not exist

                File.Copy(template, destfile, true);

                // Search for HTML in XML and convert it into altChunks
                if (allowHTML)
                {
                    using (WordprocessingDocument docWord = WordprocessingDocument.Open(destfile, true))
                    {
                        int altChunkId = 1;
                        MainDocumentPart mainPart = docWord.MainDocumentPart;

                        foreach (XmlNode node in data.SelectNodes("//text()"))
                        {
                            if (node.Value.IndexOf('<') == -1)
                                continue;

                            AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, $"htmlChunk{altChunkId}");

                            using (Stream chunkStream = chunk.GetStream(FileMode.Create, FileAccess.Write))
                            using (StreamWriter stringStream = new StreamWriter(chunkStream))
                                stringStream.Write($"<html>{node.Value}</html>");

                            AltChunk altChunk = new AltChunk { Id = $"htmlChunk{altChunkId}" };
                            node.Value = $"<w:altChunk r:id=\"htmlChunk{altChunkId}\"/>";

                            ++altChunkId;
                        }

                        mainPart.Document.Save();
                    }
                }

                FileStream destfileStream = File.Open(destfile, FileMode.Open);
                using (ZipArchive zipArchive = new ZipArchive(destfileStream, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry zipFile;
                    String[] zipFiles = new String[] { @"word/document.xml", @"word/header1.xml", @"word/header2.xml", @"word/header3.xml", @"word/header4.xml", @"word/footer1.xml", @"word/footer2.xml", @"word/footer3.xml", @"word/footer4.xml" };

                    for (int i = 0; i < zipFiles.Length; ++i)
                    {
                        zipFile = zipArchive.GetEntry(zipFiles[i]);
                        if (zipFile == null)
                            continue;

                        StreamReader reader = new StreamReader(zipFile.Open());
                        string body = reader.ReadToEnd();
                        reader.Close();
                        zipFile.Delete();

                        body = Cleanup(body);

                        // Short tags syntax @[0-9]+ convert to @@v[0-9]+
                        if (shortTags)
                            Regex.Replace(body, @"@(\d+)", "@@v$1", RegexOptions.Compiled);
                        /*
                        if (shortTags == true)
                        {
                            body = body.Replace(@"@", @"@@v");
                            // Clean up document.xml: Microsoft Word inserts spurious tags between @@ and <tag> that prevent proper @@<tag> identification.
                            foreach (Match match in new Regex(@"@@v<\/w:t><\/w:r>.*?<w:t>([0-9]+)", RegexOptions.Compiled).Matches(body))
                            {
                                body = body.Replace(match.Value, @"@@v" + match.Groups[1].Value);
                            }
                        }
                        */
                        int limit = 0;
                        while (PLACEHOLDER.IsMatch(body) && limit < 10)
                        {
                            body = Fill(body, data.DocumentElement, novalue);

                            // Remove [hidden]
                            int h = body.IndexOf("[hidden]"), s, e;
                            while (h != -1)
                            {
                                s = body.LastIndexOf("<w:tr ", h);
                                e = body.LastIndexOf("</w:tr>", h);
                                if (s == -1 || (s != -1 && s < e))
                                { // [hidden] not wrapped inside <w:tr></w:tr> 
                                    s = body.LastIndexOf("<w:p ", h);
                                    e = body.LastIndexOf("</w:p>", h);
                                    if (s == -1 || (s != -1 && s < e)) // [hidden] not wrapped inside <w:p></w:p> just remove [hidden]
                                        body = body.Remove(h, 8);
                                    else // [hidden] wrapped inside <w:p></w:p> remove whole row
                                        body = body.Remove(s, body.IndexOf("</w:p>", h) - s + 6);
                                }
                                else // [hidden] wrapped inside <w:tr></w:tr> remove whole row
                                    body = body.Remove(s, body.IndexOf("</w:tr>", h) - s + 7);
                                h = body.IndexOf("[hidden]");
                            }
                            ++limit;
                        }

                        zipFile = zipArchive.CreateEntry(zipFiles[i]);
                        StreamWriter writer = new StreamWriter(zipFile.Open());
                        writer.Write(body);
                        writer.Flush();
                        writer.Close();
                    }

                    foreach (ZipArchiveEntry entry in zipArchive.Entries)
                        if (Regex.Match(entry.Name, @"^image\d+").Success)
                        {
                            XmlNodeList items = data.GetElementsByTagName(entry.Name[..entry.Name.IndexOf('.')]);
                            if (items.Count > 0 && File.Exists(items[0].InnerText))
                            {
                                // Console.WriteLine($"Replace {entry.Name} with {items[0].InnerText}");
                                try
                                {
                                    zipArchive.CreateEntryFromFile(items[0].InnerText, entry.FullName);
                                    entry.Delete();
                                }
                                catch
                                {
                                }
                            }
                        }
                }
                destfileStream.Close();
            }
            catch (SystemException e)
            {
                if (e.GetType().Name != "InvalidOperationException") // Collection was modified; enumeration operation may not execute.
                    return String.Format("{0}: {1}", e.GetType().Name, e.Message);
            }
        pdf:
            return CreatePDF(destfile, pdf);
        }
        private static string CreatePDF(string destfile, bool pdf)
        {
            if (pdf)
            {
                // dotnet add package Spire.Doc
                Spire.Doc.Document pdfDoc = new Spire.Doc.Document();
                pdfDoc.LoadFromFile(destfile);
                Spire.Doc.ToPdfParameterList parms = new Spire.Doc.ToPdfParameterList()
                {
                    IsEmbeddedAllFonts = true
                };
                destfile = destfile.Replace(".docx", ".pdf");
                pdfDoc.SaveToFile(destfile, parms);
                pdfDoc.Close();
            }
            return destfile;
        }

        // https://json.org/json-it.html
        private static string _json = "";
        private static readonly Regex REGEX_NAME = new Regex(@"^\s*\{\s*""(?<name>[a-z0-9_]+?)""\s*:\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex REGEX_REST = new Regex(@"^\s*\[\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex REGEX_NAME_REST = new Regex(@"^""(?<name>[a-z0-9_]+?)""\s*:\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex REGEX_VALUE = new Regex(@"^(?<value>true|false|null|-?(?:0|[1-9])[0-9]*(?:\.[0-9]+)?(?:e[\-+]?[0-9]+)?|""(?:\\""|.)*?"")\s*,?\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex REGEX_FINAL = new Regex(@"^(?:[\]}]\s*,?\s*)(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static string Json2Xml(string json, string name = "root")
        {
            if (json == "")
                return "";

            Match match = REGEX_NAME.Match(json);
            if (match.Success)
            {
                if (name != "")
                    return String.Format("<{0}>{1}</{0}>", name, Json2Xml(match.Groups["rest"].Value, match.Groups["name"].Value)) + Json2Xml(_json, "");
                return Json2Xml(match.Groups["rest"].Value, match.Groups["name"].Value) + Json2Xml(_json, "");
            }

            match = REGEX_REST.Match(json);
            if (match.Success)
            {
                if (name != "")
                    return String.Format("<{0}>{1}</{0}>", name, Json2Xml(match.Groups["rest"].Value, match.Groups["name"].Value)) + Json2Xml(_json, "");
                return Json2Xml(match.Groups["rest"].Value, match.Groups["name"].Value) + Json2Xml(_json, "");
            }

            match = REGEX_NAME_REST.Match(json);
            if (match.Success)
                return Json2Xml(match.Groups["rest"].Value, match.Groups["name"].Value);

            match = REGEX_VALUE.Match(json);
            if (match.Success)
            {
                if (name == "" || name == "root")
                    name = "value";
                string value = match.Groups["value"].Value.Trim('\"');
                if (value == "null")
                    value = "";
                return String.Format("<{0}>{1}</{0}>", name, SecurityElement.Escape(value)) + Json2Xml(match.Groups["rest"].Value);
            }

            match = REGEX_FINAL.Match(json);
            if (match.Success)
            {
                _json = match.Groups["rest"].Value;
                return "";
            }

            throw new System.Data.SyntaxErrorException("Invalid JSON syntax");
        }
        static void Main(string[] args)
        {
            string template = @".\template.docx", data = @".\data.xml", destfile = @"document.docx", novalue = @"***", mime = "application/xml", args_path = "";
            bool overwrite = false, pdf = false, shorttags = false, allowhtml = false;

            try
            {
                for (int i = 0; i < args.Length; ++i)
                {
                    if ((args[i] == "--template" || args[i] == "-t") && args[i + 1].EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase)) // Case sensitive
                        template = args[++i];
                    else if (args[i] == "--xml" || args[i] == "-x")
                    {
                        mime = "application/xml";
                        data = args[++i];
                    }
                    else if (args[i] == "--json")
                    {
                        mime = "application/json";
                        data = args[++i];
                    }
                    else if ((args[i] == "--destfile" || args[i] == "-d") && args[i + 1].EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase)) // Case sensitive
                        destfile = args[++i];
                    else if (args[i] == "--overwrite" || args[i] == "-o")
                        overwrite = true;
                    else if (args[i] == "--novalue")
                        novalue = i + 1 < args.Length ? args[++i] : "";
                    else if (args[i] == "--shorttags")
                        shorttags = true;
                    else if (args[i] == "--pdf")
                        pdf = true;
                    else if (args[i] == "--allowhtml")
                        allowhtml = true;
                    else
                    {
                        args_path = args[i];
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(args_path);
                        template = xmlDoc.SelectSingleNode(@"//template")?.InnerText;
                        data = xmlDoc.SelectSingleNode(@"//data")?.InnerText;
                        mime = xmlDoc.SelectSingleNode(@"//mime")?.InnerText;
                        destfile = xmlDoc.SelectSingleNode(@"//destfile")?.InnerText;
                        novalue = xmlDoc.SelectSingleNode(@"//novalue")?.InnerText;
                        overwrite = xmlDoc.SelectSingleNode(@"//overwrite")?.InnerText == "true";
                        shorttags = xmlDoc.SelectSingleNode(@"//shorttags")?.InnerText == "true";
                        pdf = xmlDoc.SelectSingleNode(@"//pdf")?.InnerText == "true";
                        allowhtml = xmlDoc.SelectSingleNode(@"//allowhtml")?.InnerText == "true";
                        xmlDoc = null;
                    }
                }
            }
            catch (SystemException e)
            {
                if (e.HResult == -2146232000)
                    Console.WriteLine($"{e.Message} [${args_path}]");
                else
                    Console.WriteLine(@"usage: filldocx [<args_path>] --template <path> (--xml|--json) (<path>|<url>|<raw>) --destfile <path> [--pdf] [--overwrite] [--shorttags] [--allowhtml] [--novalue <string>]");
                return;
            }
            Console.WriteLine(FillDOCX(template, mime, data, destfile, novalue, overwrite, pdf, shorttags, allowhtml));
        }
    }
}
