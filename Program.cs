//#define PDF

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO.Compression;
#if (PDF)
using SautinSoft.Document; // dotnet add package SautinSoft.Document
#endif

namespace FillDOCX
{
    class Program
    {
        static Regex placeholder = new Regex(@"@@([a-z]\w*)\.?([a-z]\w*)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static string Fill(string template, XmlElement data, string novalue = "***", int level = 1)
        {
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
                    if (nodes[0].InnerText != nodes[0].LastChild.InnerText) // Has children
                    {
                        // Repeating placeholders must be placed inside tables, the subtemplate matches the row containing the placeholder
                        subtemplate = new Regex(@"\<w\:tr(?:(?!\<w\:tr).)*?@@" + tag + @".*?\<\/w\:tr\>", RegexOptions.Compiled).Match(template).Value;
                        foreach (XmlElement node in nodes)
                            value += Fill(subtemplate, node, novalue, level + 1);
                    }
                    else
                        value = nodes[0].InnerXml;
                }
                template = template.Replace(subtemplate, value);
            });

            return template;
        }

        public static string FillDOCX(string template, string xml, string destfile, string novalue, bool pdf = false, bool shortTags = false)
        {
            XmlDocument data = new XmlDocument();
            data.PreserveWhitespace = true;
            try { data.Load(xml); } catch { data.LoadXml(xml); }

            try
            {
                if (template == destfile)
                    throw new ArgumentException("Template cannot be equal to destination");

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
                            foreach (Match match in new Regex(@"@@v</w:t></w:r>.*?<w:t>([0-9]+)", RegexOptions.Compiled).Matches(body))
                            {
                                body = body.Replace(match.Value, @"@@v" + match.Groups[1].Value);
                            }
                        }
                        else
                            // Clean up document.xml: Microsoft Word inserts spurious tags between @@ and <tag> that prevent proper @@<tag> identification.
                            foreach (Match match in new Regex(@"(@@|@@([a-z]\w*)\.)</w:t></w:r>.*?<w:t>", RegexOptions.Compiled).Matches(body))
                            {
                                body = body.Replace(match.Value, match.Groups[1].Value);
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

#if PDF               
                if (pdf)
                {
                    DocumentCore dc = DocumentCore.Load(destfile);
                    destfile = destfile.Replace(".docx", ".pdf");
                    dc.Save(destfile);
                }
#endif
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
            bool pdf = false, shortTags = false;

            for (int i = 0; i < args.Length; ++i)
            {
                if ((args[i] == "--template" || args[i] == "-t") && args[i + 1].EndsWith(".docx"))
                    template = args[++i];
                else if (args[i] == "--xml" || args[i] == "-x")
                    xml = args[++i];
                else if ((args[i] == "--destfile" || args[i] == "-d") && args[i + 1].EndsWith(".docx"))
                    destfile = args[++i];
                else if (args[i] == "--novalue")
                    novalue = args[++i];
                else if (args[i] == "--shorttags")
                    shortTags = true;
#if PDF
                else if (args[i] == "--pdf")
                    pdf = true;
#endif                    
                else
                    Console.Write("usage: filldocx --template <path> --xml {<path>|<url>|<raw>} --destfile <path> [--pdf] [--shorttags] [--novalue <string>]");
            }

            Console.WriteLine(FillDOCX(template, xml, destfile, novalue, pdf, shortTags));
        }
    }
}
