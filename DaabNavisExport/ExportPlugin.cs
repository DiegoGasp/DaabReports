using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using DaabNavisExport.Parsing;
using DaabNavisExport.Utilities;

namespace DaabNavisExport
{
    [Plugin(
        "DaabNavisExport",
        "DAAB",
        DisplayName = "Daab Navis Export",
        ToolTip = "Exports Navisworks viewpoints and comments to Daab Reports format")]
    public class ExportPlugin : AddInPlugin
    {
        private const string DbFolderName = "DB";
        private const string ImagesFolderName = "Images";

        public override int Execute(params string[] parameters)
        {
            try
            {
                if (Application.ActiveDocument == null)
                {
                    MessageBox.Show("No active document open.", "Daab Navis Export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }

                var document = Application.ActiveDocument;
                var outputDirectory = ResolveOutputDirectory(parameters);
                var exportContext = BuildExportContext(document, outputDirectory);

                ExportViewpointsToXml(document, exportContext);

                var parser = new NavisworksXmlParser();
                var parseResult = parser.Process(exportContext.XmlPath);
                parser.WriteOutputs(parseResult, exportContext.DbDirectory);

                ExportViewpointImages(exportContext, parseResult.Rows);

                MessageBox.Show(
                    $"Export complete.\nProject folder: {exportContext.ProjectDirectory}\nXML: {exportContext.XmlPath}\nCSV: {Path.Combine(exportContext.DbDirectory, NavisworksXmlParser.CsvFileName)}",
                    "Daab Navis Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Daab Navis Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }

        private static ExportContext BuildExportContext(Document document, string outputDirectory)
        {
            Directory.CreateDirectory(outputDirectory);

            var projectFolderName = ResolveProjectFolderName(document);
            var projectDirectory = Path.Combine(outputDirectory, projectFolderName);
            var dbDirectory = Path.Combine(projectDirectory, DbFolderName);
            var imagesDirectory = Path.Combine(projectDirectory, ImagesFolderName);

            Directory.CreateDirectory(projectDirectory);
            Directory.CreateDirectory(dbDirectory);
            Directory.CreateDirectory(imagesDirectory);

            var xmlFile = Path.Combine(dbDirectory, "DB.xml");

            return new ExportContext(document, outputDirectory, projectDirectory, dbDirectory, imagesDirectory, xmlFile);
        }

        private static string ResolveProjectFolderName(Document document)
        {
            var sourceName = document.FileName;
            if (!string.IsNullOrWhiteSpace(sourceName))
            {
                var stem = Path.GetFileNameWithoutExtension(sourceName);
                var sanitized = PathSanitizer.ToSafeFileName(stem);
                if (!string.IsNullOrWhiteSpace(sanitized))
                {
                    return sanitized;
                }
            }

            return $"Navisworks_{DateTime.Now:yyyyMMdd_HHmmss}";
        }

        private static string ResolveOutputDirectory(IReadOnlyList<string> parameters)
        {
            var explicitPath = parameters?.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
            if (!string.IsNullOrEmpty(explicitPath))
            {
                return explicitPath;
            }

            var navisTemp = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(navisTemp, "DaabNavisExport");
        }

        private static void ExportViewpointsToXml(Document document, ExportContext context)
        {
            context.ViewSequence.Clear();

            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = true,
                IndentChars = "  ",
                NewLineOnAttributes = false
            };

            Directory.CreateDirectory(Path.GetDirectoryName(context.XmlPath)!);

            using var writer = XmlWriter.Create(context.XmlPath, settings);
            writer.WriteStartDocument();
            writer.WriteStartElement("exchange");
            writer.WriteAttributeString("units", document.Units.ToString());
            var sourcePath = document.FileName ?? string.Empty;
            var fileName = string.IsNullOrEmpty(sourcePath) ? string.Empty : Path.GetFileName(sourcePath) ?? string.Empty;
            writer.WriteAttributeString("filename", fileName);
            writer.WriteAttributeString("filepath", sourcePath);

            writer.WriteStartElement("viewpoints");
            foreach (SavedItem item in document.SavedViewpoints.RootItems)
            {
                WriteSavedItem(writer, item, context);
            }
            writer.WriteEndElement(); // viewpoints

            writer.WriteEndElement(); // exchange
            writer.WriteEndDocument();
        }

        private static void WriteSavedItem(XmlWriter writer, SavedItem item, ExportContext context)
        {
            switch (item)
            {
                case GroupItem folder:
                    WriteFolder(writer, folder, context);
                    break;
                case SavedViewpoint viewpoint:
                    WriteView(writer, viewpoint, context);
                    break;
            }
        }

        private static void WriteFolder(XmlWriter writer, GroupItem folder, ExportContext context)
        {
            writer.WriteStartElement("viewfolder");
            writer.WriteAttributeString("name", folder.DisplayName ?? string.Empty);
            writer.WriteAttributeString("guid", folder.Guid.ToString());

            foreach (SavedItem child in folder.Children)
            {
                WriteSavedItem(writer, child, context);
            }

            writer.WriteEndElement();
        }

        private static void WriteView(XmlWriter writer, SavedViewpoint viewpoint, ExportContext context)
        {
            context.ViewSequence.Add(viewpoint);

            writer.WriteStartElement("view");
            writer.WriteAttributeString("name", viewpoint.DisplayName ?? string.Empty);
            writer.WriteAttributeString("guid", viewpoint.Guid.ToString());

            WriteComments(writer, viewpoint);

            writer.WriteEndElement();
        }

        private static void WriteComments(XmlWriter writer, SavedViewpoint viewpoint)
        {
            try
            {
                var commentsProperty = viewpoint.GetType().GetProperty("Comments");
                if (commentsProperty == null)
                {
                    return;
                }

                if (commentsProperty.GetValue(viewpoint) is not IEnumerable comments)
                {
                    return;
                }

                var anyComments = false;
                foreach (var comment in comments)
                {
                    if (comment == null)
                    {
                        continue;
                    }

                    var commentType = comment.GetType();
                    var guid = commentType.GetProperty("Guid")?.GetValue(comment)?.ToString();
                    var status = commentType.GetProperty("Status")?.GetValue(comment)?.ToString();
                    var author = commentType.GetProperty("Author")?.GetValue(comment)?.ToString();
                    var body = commentType.GetProperty("Body")?.GetValue(comment)?.ToString();
                    var creationDate = commentType.GetProperty("CreationDate")?.GetValue(comment);

                    if (!anyComments)
                    {
                        writer.WriteStartElement("comments");
                        anyComments = true;
                    }

                    writer.WriteStartElement("comment");
                    if (!string.IsNullOrWhiteSpace(guid))
                    {
                        writer.WriteAttributeString("id", guid);
                    }

                    if (!string.IsNullOrWhiteSpace(status))
                    {
                        writer.WriteAttributeString("status", status);
                    }

                    if (!string.IsNullOrWhiteSpace(author))
                    {
                        writer.WriteElementString("user", author);
                    }

                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        writer.WriteElementString("body", body);
                    }

                    WriteCreatedDate(writer, creationDate);
                    writer.WriteEndElement();
                }

                if (anyComments)
                {
                    writer.WriteEndElement();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to reflect Navisworks comments: {ex.Message}");
            }
        }

        private static void WriteCreatedDate(XmlWriter writer, object? created)
        {
            if (created is not DateTime createdDate)
            {
                return;
            }

            if (createdDate.Year < 1900)
            {
                return;
            }

            writer.WriteStartElement("createddate");
            writer.WriteStartElement("date");
            writer.WriteAttributeString("year", createdDate.Year.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("month", createdDate.Month.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("day", createdDate.Day.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void ExportViewpointImages(ExportContext context, IEnumerable<IReadOnlyList<string?>> rows)
        {
            if (context.ViewSequence.Count == 0)
            {
                return;
            }

            Directory.CreateDirectory(context.ImagesDirectory);

            var imageAssignments = new Dictionary<Guid, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in rows)
            {
                if (row.Count <= 10)
                {
                    continue;
                }

                var guidText = row[4];
                var imagePath = row[10];
                if (string.IsNullOrWhiteSpace(guidText) || string.IsNullOrWhiteSpace(imagePath))
                {
                    continue;
                }

                if (!Guid.TryParse(guidText, out var guid))
                {
                    continue;
                }

                if (!imageAssignments.ContainsKey(guid))
                {
                    imageAssignments.Add(guid, imagePath);
                }
            }

            foreach (var viewpoint in context.ViewSequence)
            {
                if (!imageAssignments.TryGetValue(viewpoint.Guid, out var imageFile))
                {
                    continue;
                }

                var targetPath = Path.Combine(context.ImagesDirectory, imageFile);
                using var bitmap = viewpoint.GenerateThumbnail(new Size(800, 450));
                if (bitmap == null)
                {
                    continue;
                }

                bitmap.Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }
    }
}
