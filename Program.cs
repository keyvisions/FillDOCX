using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO.Compression;
using Spire.Doc; // dotnet add package FreeSpire.Doc (Free PDF limited to 3 pages)
using System.Security.AccessControl;
using System.Security;
using System.Data;

namespace FillDOCX
{
    class Program
    {
        static Regex placeholder = new Regex(@"@@([a-z]\w*)\.?([a-z]\w*)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static string Fill(string template, XmlElement data, string novalue = "***", int level = 1)
        {
            if (data.Attributes.GetNamedItem("hidden") != null)
                return "";

            List<string> tags = new List<string>();
            foreach (Match match in placeholder.Matches(template))
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
                            subtemplate = new Regex(@"<w:t>@@" + Regex.Escape(tag) + @".*?<\/w:t>", RegexOptions.Compiled).Match(template).Value;
                            if (subtemplate == "")
                                return;
                            value += Fill(subtemplate, (XmlElement)nodes[0], novalue, level + 1);
                        }
                        else
                        {
                            foreach (XmlElement node in nodes)
                                value += Fill(subtemplate, node, novalue, level + 1);
                            if (value.Contains("[hidden]"))
                            { // Remove whole table
                                subtemplate = new Regex(@"<w:tbl>(?:(?!<w:tbl>).)*?" + Regex.Escape(subtemplate) + @".*?<\/w:tbl>", RegexOptions.Compiled).Match(template).Value;
                                value = "";
                            }
                        }
                    }
                    else if (nodes[0].Attributes.GetNamedItem("hidden") != null)
                        value = "";
                    else
                        value = nodes[0].InnerXml;
                }
                if (Regex.Match(tag, @"^image\d+").Success)
                    value = "";

                if (value == "[hidden]")
                    subtemplate = new Regex(@"<w:tbl>(?:(?!<w:tbl>).)*?" + Regex.Escape(subtemplate) + @".*?<\/w:tbl>", RegexOptions.Compiled).Match(template).Value;

                if (subtemplate != "")
                    template = template.Replace(subtemplate, value);
            });
            return template;
        }

        public static string FillDOCX(string template, string mime, string txt, string destfile, string novalue, bool overwrite = false, bool pdf = false, bool shortTags = false)
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
                            StreamReader sr = new StreamReader(txt);
                            txt = sr.ReadToEnd();
                            sr.Dispose();
                        }
#if DEBUG
                        Console.Write(json2xml(txt));
#endif

                        data.LoadXml(json2xml(txt));
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
                    throw new ArgumentException("Template cannot be equal to destination");

                if (File.Exists(destfile) && !overwrite)
                    goto pdf; // return destfile;

                if (Path.GetDirectoryName(destfile) != "")
                    Directory.CreateDirectory(Path.GetDirectoryName(destfile)); // Create directory if it does not exist

                File.Copy(template, destfile, true);

                using (ZipArchive zipArchive = new ZipArchive(File.Open(destfile, FileMode.Open), ZipArchiveMode.Update))
                {
                    ZipArchiveEntry zipFile;
                    String[] zipFiles = new String[] { @"word/document.xml", @"word/header1.xml", @"word/header2.xml", @"word/header3.xml", @"word/footer1.xml", @"word/footer2.xml", @"word/footer3.xml" };

                    for (int i = 0; i < zipFiles.Length; ++i)
                    {
                        zipFile = zipArchive.GetEntry(zipFiles[i]);
                        if (zipFile == null)
                            continue;

                        StreamReader reader = new StreamReader(zipFile.Open());
                        string body = reader.ReadToEnd();
                        reader.Close();
                        zipFile.Delete();

                        // Short tags syntax @[0-9]+ convert to @@v[0-9]+
                        if (shortTags == true)
                        {
                            body = body.Replace(@"@", @"@@v");
                            // Clean up document.xml: Microsoft Word inserts spurious tags between @@ and <tag> that prevent proper @@<tag> identification.
                            foreach (Match match in new Regex(@"@@v<\/w:t><\/w:r>.*?<w:t>([0-9]+)", RegexOptions.Compiled).Matches(body))
                            {
                                body = body.Replace(match.Value, @"@@v" + match.Groups[1].Value);
                            }
                        }
                        else
                        {
                            // Clean up document.xml: Microsoft Word inserts spurious tags between @@ and <tag> that prevent proper @@<tag> identification.
                            bool replace;
                            do
                            {
                                MatchCollection matches = new Regex(@"(@@\w*?\.?\w*?)<\/w:t><\/w:r>.*?<w:t( .*?>|>)(.)", RegexOptions.Compiled).Matches(body);

                                replace = false;
                                foreach (Match match in matches)
                                    if (System.Char.IsLetterOrDigit(match.Groups[3].Value, 0) || match.Groups[3].Value == @".")
                                    {
                                        body = body.Replace(match.Value, match.Groups[1].Value + match.Groups[3].Value);
                                        replace = true;
                                    }
                            } while (replace);
                        }

                        int limit = 0;
                        while (placeholder.IsMatch(body) && limit < 5)
                        {
                            body = Fill(body, data.DocumentElement, novalue);
                            ++limit;
                        }

                        zipFile = zipArchive.CreateEntry(zipFiles[i]);
                        StreamWriter writer = new StreamWriter(zipFile.Open());
                        writer.Write(body.Replace("\u000B", @"<w:br/>"));
                        writer.Flush();
                        writer.Close();
                    }

                    foreach (ZipArchiveEntry entry in zipArchive.Entries)
                        if (Regex.Match(entry.Name, @"^image\d+").Success)
                        {
                            XmlNodeList items = data.GetElementsByTagName(entry.Name.Substring(0, entry.Name.IndexOf('.')));
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
            }
            catch (SystemException e)
            {
                if (e.GetType().Name != "InvalidOperationException") // Collection was modified; enumeration operation may not execute.
                    return String.Format("{0}: {1}", e.GetType().Name, e.Message);
            }
        pdf:
            CreatePDF(destfile, pdf);
            return destfile;
        }
        private static void CreatePDF(string destfile, bool pdf)
        {
            if (pdf)
            {
                // dotnet add package Spire.Doc
                Document dc = new Document();
                dc.LoadFromFile(destfile);
                ToPdfParameterList parms = new ToPdfParameterList()
                {
                    IsEmbeddedAllFonts = true
                };
                destfile = destfile.Replace(".docx", ".pdf");
                dc.SaveToFile(destfile, parms);
            }
        }

        // https://json.org/json-it.html
        private static string _json = "";
        private static string json2xml(string json, string name = "root")
        {
            if (json == "")
                return "";

            Regex regex = new Regex(@"^\s*\{\s*""(?<name>[a-z0-9_]+?)""\s*:\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            Match match = regex.Match(json);
            if (match.Success)
            {
                if (name != "")
                    return String.Format("<{0}>{1}</{0}>", name, json2xml(match.Groups["rest"].Value, match.Groups["name"].Value)) + json2xml(_json, "");
                return json2xml(match.Groups["rest"].Value, match.Groups["name"].Value) + json2xml(_json, "");
            }

            regex = new Regex(@"^\s*\[\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            match = regex.Match(json);
            if (match.Success)
            {
                if (name != "")
                    return String.Format("<{0}>{1}</{0}>", name, json2xml(match.Groups["rest"].Value, match.Groups["name"].Value)) + json2xml(_json, "");
                return json2xml(match.Groups["rest"].Value, match.Groups["name"].Value) + json2xml(_json, "");
            }

            regex = new Regex(@"^""(?<name>[a-z0-9_]+?)""\s*:\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            match = regex.Match(json);
            if (match.Success)
                return json2xml(match.Groups["rest"].Value, match.Groups["name"].Value);

            regex = new Regex(@"^(?<value>true|false|null|-?(?:0|[1-9])[0-9]*(?:\.[0-9]+)?(?:e[\-+]?[0-9]+)?|""(?:\\""|.)*?"")\s*,?\s*(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            match = regex.Match(json);
            if (match.Success)
            {
                if (name == "" || name == "root")
                    name = "value";
                string value = match.Groups["value"].Value.Trim('\"');
                if (value == "null")
                    value = "";
                return String.Format("<{0}>{1}</{0}>", name, SecurityElement.Escape(value)) + json2xml(match.Groups["rest"].Value);
            }

            regex = new Regex(@"^(?:[\]}]\s*,?\s*)(?<rest>[\s\S]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            match = regex.Match(json);
            if (match.Success)
            {
                _json = match.Groups["rest"].Value;
                return "";
            }

            throw new SyntaxErrorException("Invalid JSON syntax");
        }
        static void Main(string[] args)
        {
            string template = @".\template.docx", data = @".\data.xml", destfile = @"document.docx", novalue = @"***", mime = "application/xml";
            bool overwrite = false, pdf = false, shortTags = false;

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
                    shortTags = true;
                else if (args[i] == "--pdf")
                    pdf = true;
                else
                {
                    Console.WriteLine("usage: filldocx --template <path> --{xml|json} {<path>|<url>|<raw>} --destfile <path> [--pdf] [--overwrite] [--shorttags] [--novalue <string>]");
                    return;
                }
            }

            Console.WriteLine(FillDOCX(template, mime, data, destfile, novalue, overwrite, pdf, shortTags));
        }
    }
}
