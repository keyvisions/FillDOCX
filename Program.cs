using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO.Compression;
using Spire.Doc; // dotnet add package FreeSpire.Doc (Free PDF limited to 3 pages)

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
                string subtemplate = level == 1 ? $"@@{tag}" : $"@@{data.Name}.{tag}", value = novalue;
                if (nodes.Count > 0)
                {
                    if (nodes[0].HasChildNodes && level == 1)
                    {
                        // Repeating placeholders MUST be placed inside tables, the subtemplate matches the row containing the placeholder
                        subtemplate = new Regex(@"<w:tr (?:(?!<w:tr ).)*?@@" + tag + @".*?<\/w:tr>", RegexOptions.Compiled).Match(template).Value;
                        if (subtemplate == "")
                        {
                            subtemplate = new Regex(@"<w:t>@@" + tag + @".*?<\/w:t>", RegexOptions.Compiled).Match(template).Value;
                            if (subtemplate == "")
                                return;
                            value += Fill(subtemplate, (XmlElement)nodes[0], novalue, level + 1);
                        }
                        else
                        {
                            foreach (XmlElement node in nodes)
                                value += Fill(subtemplate, node, novalue, level + 1);
                            if (value.Contains("[hidden]"))
                                value = "";
                        }
                    }
                    else if (nodes[0].Attributes.GetNamedItem("hidden") != null)
                        value = "";
                    else
                        value = nodes[0].InnerXml;
                }
                template = template.Replace(subtemplate, value);
            });
            return template;
        }

        public static string FillDOCX(string template, string xml, string destfile, string novalue, bool overwrite = false, bool pdf = false, bool shortTags = false)
        {
            XmlDocument data = new XmlDocument();
            data.PreserveWhitespace = true;
            try { data.Load(xml); } catch { data.LoadXml(xml); }

            try
            {
                if (template == destfile)
                    throw new ArgumentException("Template cannot be equal to destination");

                if (File.Exists(destfile) && !overwrite)
                    return destfile;

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
                }

                if (pdf)
                {
                    // dotnet add package Spire.Doc
                    Document dc = new Document();
                    dc.LoadFromFile(destfile);
                    destfile = destfile.Replace(".docx", ".pdf");
                    dc.SaveToFile(destfile, FileFormat.PDF);
                }

                return destfile;
            }
            catch (SystemException e)
            {
                return String.Format("{0}: {1}", e.GetType().Name, e.Message);
            }
        }
        static void Main(string[] args)
        {
            string template = @".\order.docx", xml = @"data.xml", destfile = @"document.docx", novalue = @"***";
            bool overwrite = true, pdf = true, shortTags = false;

            for (int i = 0; i < args.Length; ++i)
            {
                if ((args[i] == "--template" || args[i] == "-t") && args[i + 1].EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase)) // Case sensitive
                    template = args[++i];
                else if (args[i] == "--xml" || args[i] == "-x")
                    xml = args[++i];
                else if ((args[i] == "--destfile" || args[i] == "-d") && args[i + 1].EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase)) // Case sensitive
                    destfile = args[++i];
                else if (args[i] == "--overwrite" || args[i] == "-o")
                    overwrite = true;
                else if (args[i] == "--novalue")
                    novalue = args[++i];
                else if (args[i] == "--shorttags")
                    shortTags = true;
                else if (args[i] == "--pdf")
                    pdf = true;
                else
                    Console.Write("usage: filldocx --template <path> --xml {<path>|<url>|<raw>} --destfile <path> [--pdf] [--overwrite] [--shorttags] [--novalue <string>]");
            }

            Console.WriteLine(FillDOCX(template, xml, destfile, novalue, overwrite, pdf, shortTags));
        }
    }
}
