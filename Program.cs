using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO.Compression;
using Spire.Doc; // https://www.e-iceblue.com/Introduce/spire-office-for-net-free.html
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Web;
using System.Text.Json;
using System.Net.Http;
using System.Threading.Tasks;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;

namespace FillDOCX
{
	class Program
	{
		private static readonly Regex PLACEHOLDER = new Regex(@"@@(\w+)(?>\.(\w+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase);
		private static ushort _novalue = 0x0; // Keep track of novalue
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
				_novalue |= 0x1;

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
					{
						value = "";
						_novalue &= 0x2;
					}
					else
					{
						value = nodes[0].InnerXml;
						_novalue &= 0x2;
					}
				}
				if (Regex.Match(tag, @"^image\d+").Success)
					value = "";

				if (subtemplate != "")
				{
					if (value.IndexOf("altChunk") != -1)
						value = HttpUtility.HtmlDecode(value);
					value = Regex.Replace(value, @"&amp;(lt;|gt;|quot;|apos;)", "&$1", RegexOptions.Multiline | RegexOptions.Compiled);
					value = value.Replace("\\n", "<w:br/>");
					value = value.Replace("\\/", "/");
					template = template.Replace(subtemplate, value);
				}

				if ((_novalue & 0x1) == 0x1)
					_novalue = 0x2;
			});
			return template;
		}
		private static readonly Regex USELESS = new Regex(@"</w:t></w:r><[\s\S]*?(<w:t>|<w:t [\s\S]*?>)(?<whitespace>.{1})", RegexOptions.Compiled | RegexOptions.IgnoreCase);
		private static string Cleanup(string body)
		{

			XmlDocument xmlDoc = new XmlDocument();
			xmlDoc.LoadXml(body);

			int s = -2, sc;
			MatchCollection matches = PLACEHOLDER.Matches(xmlDoc.InnerText);
			foreach (Match match in matches.Cast<Match>())
			{
				string tag = match.Value;

				sc = body.IndexOf("@</w:t>", s + 2); // Account for degenerate case, stand alone @ at end of a w:t run
				s = body.IndexOf("@@", s + 2);
				if (sc != -1 && sc < s)
					s = sc;

				int i = 0; // i to exit potential infinite loop
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
		private static async Task<string> FillDOCX(string template, string mime, string txt, string destfile, string novalue, bool overwrite = false, bool pdf = false, bool shortTags = false, bool allowHTML = false, bool ignoreincomplete = false)
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
						if (!Regex.IsMatch(txt, @"\s*[\[{]") && txt.IndexOfAny(Path.GetInvalidPathChars()) == -1)
						{
							if (txt.StartsWith("http"))
							{
								txt = await FetchDataAsync(txt);
							}
							else
							{
								using StreamReader sr = new StreamReader(txt);
								txt = sr.ReadToEnd();
							}
						}
#if DEBUG
						Console.Write(JsonToXml(txt));
#endif
						data.LoadXml(JsonToXml(txt));
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
					using WordprocessingDocument docWord = WordprocessingDocument.Open(destfile, true);
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

						mainPart.Document.Save();

						++altChunkId;
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

						// Short tags syntax @[0-9]+ convert to @@v[0-9]+
						if (shortTags)
							foreach (Match match in new Regex(@"@(\d*)(?:<\/w:t><\/w:r>.*?<w:t>)?(\d+)", RegexOptions.Compiled).Matches(body))
								body = body.Replace(match.Value, @"@@v" + match.Groups[1].Value + match.Groups[2].Value);
						else
							body = Cleanup(body);

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

					Dictionary<ZipArchiveEntry, string> images = new Dictionary<ZipArchiveEntry, string>();
					foreach (ZipArchiveEntry entry in zipArchive.Entries)
						if (Regex.Match(entry.Name, @"^image\d+").Success)
						{
							string image = entry.Name[..entry.Name.IndexOf('.')];
							XmlNodeList items = data.SelectNodes("//" + image + "|//" + image.ToUpper());
							if (items.Count > 0 && (items[0].InnerText.StartsWith("qrcode://") || File.Exists(items[0].InnerText)))
								images.Add(entry, items[0].InnerText);
						}
					foreach (KeyValuePair<ZipArchiveEntry, string> image in images)
					{
						if (image.Value.StartsWith("qrcode://"))
						{
							string qrCodePath = Path.Combine(Path.GetTempPath(), $"{image.Key.Name}.png");

							// Generate QR code if the image file does not exist
							using QRCoder.QRCodeGenerator qrGenerator = new QRCoder.QRCodeGenerator();
							using QRCoder.QRCodeData qrCodeData = qrGenerator.CreateQrCode(image.Value.Substring(9), QRCoder.QRCodeGenerator.ECCLevel.Q);
							using QRCoder.BitmapByteQRCode qrCode = new QRCoder.BitmapByteQRCode(qrCodeData);
							using (MemoryStream ms = new MemoryStream(qrCode.GetGraphic(20)))
							using (Bitmap qrCodeImage = new Bitmap(ms))

								qrCodeImage.Save(qrCodePath, ImageFormat.Png);
							zipArchive.CreateEntryFromFile(qrCodePath, image.Key.FullName);
							File.Delete(qrCodePath); // Clean up the temporary QR code image
						}
						else
							zipArchive.CreateEntryFromFile(image.Value, image.Key.FullName);

						image.Key.Delete();
					}
				}
				destfileStream.Close();

				if ((_novalue & 0x2) == 0x2 && ignoreincomplete == false)
				{
					string incompletefile = destfile.Replace(".docx", "__.docx");
					if (File.Exists(incompletefile))
						File.Delete(incompletefile);
					File.Move(destfile, incompletefile);
					destfile = incompletefile;
				}
			}
			catch (SystemException e)
			{
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

				File.Delete(destfile.Replace(".pdf", ".docx"));
			}
			return destfile;
		}

		private static string JsonToXml(string json, string rootElement = "root")
		{
			try
			{
				using JsonDocument document = JsonDocument.Parse(json);
				XmlDocument xmlDoc = new XmlDocument();
				XmlElement root = xmlDoc.CreateElement(rootElement);
				xmlDoc.AppendChild(root);
				ConvertJsonToXml(document.RootElement, root, xmlDoc);
				return xmlDoc.OuterXml;
			}
			catch (JsonException ex)
			{
				throw new ArgumentException("Invalid JSON format", ex);
			}
		}

		private static void ConvertJsonToXml(JsonElement jsonElement, XmlElement parentElement, XmlDocument xmlDoc)
		{
			switch (jsonElement.ValueKind)
			{
				case JsonValueKind.Object:
					foreach (JsonProperty property in jsonElement.EnumerateObject())
					{
						XmlElement childElement = xmlDoc.CreateElement(property.Name);
						parentElement.AppendChild(childElement);
						ConvertJsonToXml(property.Value, childElement, xmlDoc);
					}
					break;

				case JsonValueKind.Array:
					foreach (JsonElement arrayElement in jsonElement.EnumerateArray())
					{
						XmlElement arrayItem = xmlDoc.CreateElement("Item");
						parentElement.AppendChild(arrayItem);
						ConvertJsonToXml(arrayElement, arrayItem, xmlDoc);
					}
					break;

				case JsonValueKind.String:
					parentElement.InnerText = jsonElement.GetString();
					break;

				case JsonValueKind.Number:
					parentElement.InnerText = jsonElement.GetRawText();
					break;

				case JsonValueKind.True:
				case JsonValueKind.False:
					parentElement.InnerText = jsonElement.GetBoolean().ToString();
					break;

				case JsonValueKind.Null:
					// Leave the element empty for null values
					break;

				default:
					throw new NotSupportedException($"Unsupported JSON value kind: {jsonElement.ValueKind}");
			}
		}

		private static async Task<string> FetchDataAsync(string url)
		{
			using HttpClient client = new HttpClient();
			try
			{
				return await client.GetStringAsync(url);
			}
			catch (HttpRequestException ex)
			{
				throw new Exception($"Error fetching data from URL: {url}", ex);
			}
		}

		static async Task Main(string[] args)
		{
			string template = @".\template.docx", data = @".\data.xml", destfile = @"document.docx", novalue = @"***", mime = "application/xml", args_path = "";
			bool overwrite = false, pdf = false, shorttags = false, allowhtml = false, ignoreincomplete = false;

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
					else if (args[i] == "--ignoreincomplete")
						ignoreincomplete = true;
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
						ignoreincomplete = xmlDoc.SelectSingleNode(@"//ignoreincomplete")?.InnerText == "true";
						xmlDoc = null;
					}
				}
			}
			catch (SystemException e)
			{
				if (e.HResult == -2146232000)
					Console.WriteLine($"{e.Message} [${args_path}]");
				else
					Console.WriteLine(@"usage: filldocx [<args_path>] --template <path> (--xml|--json) (<path>|<url>|<raw>) --destfile <path> [--pdf] [--overwrite] [--shorttags] [--allowhtml] [--novalue <string>] [--ignoreincomplete]");
				return;
			}

			// Await the asynchronous FillDOCX method
			string result = await FillDOCX(template, mime, data, destfile, novalue, overwrite, pdf, shorttags, allowhtml, ignoreincomplete);
			Console.WriteLine(result);
		}
	}
}
