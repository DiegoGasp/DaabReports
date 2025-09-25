using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace DaabNavisExport.Parsing
{
    internal sealed class NavisworksXmlParser
    {
        public const string CsvFileName = "navisworks_views_comments.csv";
        public const string DebugFileName = "debug.txt";
        private const string ImageFileStem = "vp";

        public ParseResult Process(string xmlPath, bool streamDebug = false)
        {
            if (!File.Exists(xmlPath))
            {
                throw new FileNotFoundException("XML file not found", xmlPath);
            }

            var rows = new List<List<string?>>();
            var debug = new List<string>();
            var seen = new HashSet<(string? Guid, string? CommentId)>();
            var viewCounter = 0;
            var imagePrefix = Path.GetFileNameWithoutExtension(CsvFileName);

            void Log(string message)
            {
                debug.Add(message);
                if (streamDebug)
                {
                    System.Diagnostics.Debug.WriteLine(message);
                }
            }

            var document = XDocument.Load(xmlPath);
            var root = document.Root ?? throw new InvalidDataException("Invalid XML: missing root");
            var viewFolders = root.Element("viewpoints")?.Elements("viewfolder") ?? Enumerable.Empty<XElement>();

            var imagePrefixBase = BuildImagePrefix();

            foreach (var folder in viewFolders)
            {
                RecurseFolder(folder, new List<string>(), rows, seen, ref viewCounter, imagePrefixBase, Log);
            }

            return new ParseResult(rows, debug);
        }

        public void WriteOutputs(ParseResult result, string outputDirectory)
        {
            Directory.CreateDirectory(outputDirectory);

            var csvPath = Path.Combine(outputDirectory, CsvFileName);
            using (var writer = new StreamWriter(csvPath, false, new UTF8Encoding(false)))
            {
                WriteCsvLine(writer, new[]
                {
                    "Category",
                    "Level",
                    "Subfolder",
                    "ViewName",
                    "GUID",
                    "CommentID",
                    "Status",
                    "User",
                    "Body",
                    "CreatedDate",
                    "ImagePath"
                });

                foreach (var row in result.Rows)
                {
                    WriteCsvLine(writer, row.Select(field => field ?? string.Empty));
                }
            }

            File.WriteAllLines(Path.Combine(outputDirectory, DebugFileName), result.DebugLines, new UTF8Encoding(false));
        }

        private static void WriteCsvLine(TextWriter writer, IEnumerable<string> fields)
        {
            var builder = new StringBuilder();
            var first = true;
            foreach (var field in fields)
            {
                if (!first)
                {
                    builder.Append(',');
                }
                else
                {
                    first = false;
                }

                builder.Append('"');
                builder.Append(field.Replace("\"", "\"\""));
                builder.Append('"');
            }

            writer.WriteLine(builder.ToString());
        }

        private static void RecurseFolder(
            XElement folder,
            List<string> path,
            ICollection<List<string?>> rows,
            ISet<(string? Guid, string? CommentId)> seen,
            ref int viewCounter,
            string imagePrefix,
            Action<string> log)
        {
            var folderName = folder.Attribute("name")?.Value ?? string.Empty;
            var newPath = new List<string>(path) { folderName };

            log($"📂 Entering folder: {string.Join(" > ", newPath.Where(p => !string.IsNullOrWhiteSpace(p)))}");

            foreach (var view in folder.Elements("view"))
            {
                viewCounter++;
                var viewName = view.Attribute("name")?.Value ?? string.Empty;
                var guid = view.Attribute("guid")?.Value ?? string.Empty;
                var imageFile = $"{imagePrefix}{viewCounter.ToString("0000", CultureInfo.InvariantCulture)}.jpg";

                log($"  👀 Found view: {viewName} (GUID={guid}) → {imageFile}");

                var commentsNode = view.Element("comments");
                if (commentsNode != null)
                {
                    foreach (var comment in commentsNode.Elements("comment"))
                    {
                        var row = BuildRow(newPath, viewName, guid, comment, imageFile, log);
                        AddRow(row, rows, seen, log);
                    }
                }
                else
                {
                    var row = BuildEmptyCommentRow(newPath, viewName, guid, imageFile);
                    AddRow(row, rows, seen, log);
                }
            }

            foreach (var child in folder.Elements("viewfolder"))
            {
                RecurseFolder(child, newPath, rows, seen, ref viewCounter, imagePrefix, log);
            }
        }

        private static string BuildImagePrefix()
        {
            var csvStem = Path.GetFileNameWithoutExtension(CsvFileName);
            return string.IsNullOrWhiteSpace(csvStem)
                ? $"{ImageFileStem}_"
                : $"{csvStem}_{ImageFileStem}";
        }

        private static List<string?> BuildRow(
            IReadOnlyList<string> path,
            string viewName,
            string guid,
            XElement comment,
            string imageFile,
            Action<string> log)
        {
            var commentId = comment.Attribute("id")?.Value;
            var status = comment.Attribute("status")?.Value;
            var user = comment.Element("user")?.Value;
            var body = comment.Element("body")?.Value;
            var created = ParseCreatedDate(comment.Element("createddate"), log);

            log($"    💬 Comment ID={commentId}, Status={status}, User={user}");

            return new List<string?>
            {
                path.ElementAtOrDefault(0),
                path.ElementAtOrDefault(1),
                path.Count > 2 ? string.Join(" > ", path.Skip(2)) : null,
                viewName,
                guid,
                commentId,
                status,
                user,
                body,
                created,
                imageFile
            };
        }

        private static List<string?> BuildEmptyCommentRow(
            IReadOnlyList<string> path,
            string viewName,
            string guid,
            string imageFile)
        {
            return new List<string?>
            {
                path.ElementAtOrDefault(0),
                path.ElementAtOrDefault(1),
                path.Count > 2 ? string.Join(" > ", path.Skip(2)) : null,
                viewName,
                guid,
                null,
                null,
                null,
                null,
                null,
                imageFile
            };
        }

        private static string? ParseCreatedDate(XElement? createdNode, Action<string> log)
        {
            try
            {
                var dateNode = createdNode?.Element("date");
                if (dateNode == null)
                {
                    return null;
                }

                var year = SafeParse(dateNode.Attribute("year")?.Value);
                var month = SafeParse(dateNode.Attribute("month")?.Value);
                var day = SafeParse(dateNode.Attribute("day")?.Value);

                if (year < 1900 || month <= 0 || day <= 0)
                {
                    return null;
                }

                var dt = new DateTime(year, month, day);
                return dt.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                log($"❌ Failed to parse createddate: {ex.Message}");
                return null;
            }
        }

        private static int SafeParse(string? value)
        {
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;
        }

        private static void AddRow(
            List<string?> row,
            ICollection<List<string?>> rows,
            ISet<(string? Guid, string? CommentId)> seen,
            Action<string> log)
        {
            var key = (row.ElementAtOrDefault(4), row.ElementAtOrDefault(5));
            if (seen.Contains(key))
            {
                log($"⚠️ Duplicate skipped: GUID={key.Item1}, CommentID={key.Item2}");
                return;
            }

            seen.Add(key);
            rows.Add(row);
        }
    }
}
