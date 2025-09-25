using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
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
using NavisApplication = Autodesk.Navisworks.Api.Application;

namespace DaabNavisExport
{
    [Plugin(
        "DaabNavisExport",
        "DAAB",
        DisplayName = "DaabReport",
        ToolTip = "Exports Navisworks viewpoints and comments to Daab Reports format")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ExportPlugin : AddInPlugin
    {
        private const string DbFolderName = "DB";
        private const string ImagesFolderName = "Images";

        public override int Execute(params string[] parameters)
        {
            try
            {
                if (NavisApplication.ActiveDocument == null)
                {
                    MessageBox.Show("No active document open.", "Daab Navis Export",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }

                var document = NavisApplication.ActiveDocument;
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
                MessageBox.Show($"Export failed: {ex.Message}", "Daab Navis Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            var rootItem = document.SavedViewpoints?.RootItem;
            if (rootItem != null)
            {
                foreach (SavedItem item in rootItem.Children)
                {
                    WriteSavedItem(writer, item, context);
                }
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

            var imageAssignments = new Dictionary<Guid, string>();
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
                var targetDirectory = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                var rendered = TryRenderViewpointImage(context.Document, viewpoint, targetPath, new Size(800, 450));
                if (!rendered)
                {
                    rendered = TryRenderViewpointImage(context.Document, viewpoint, targetPath, new Size(800, 450));
                }

                if (!rendered)
                {
                    using var bitmap = TryGenerateThumbnail(viewpoint, new Size(800, 450));
                    if (bitmap != null)
                    {
                        SaveBitmapToJpeg(bitmap, targetPath);
                        rendered = File.Exists(targetPath);
                        if (!rendered)
                        {
                            Debug.WriteLine($"Thumbnail save reported success but file not found at {targetPath}.");
                        }
                    }
                }

                if (!rendered)
                {
                    Debug.WriteLine($"Unable to produce image for viewpoint {viewpoint.DisplayName} ({viewpoint.Guid}).");
                }
            }
        }

        private static bool TryRenderViewpointImage(Document document, SavedViewpoint viewpoint, string targetPath, Size size)
        {
            try
            {
                TryApplyViewpoint(document, viewpoint);

                var activeView = document.ActiveView;
                if (activeView == null)
                {
                    return false;
                }

                var style = ImageGenerationStyle.ScenePlusOverlay;

                if (activeView.SaveToImage(targetPath, style, size.Width, size.Height))
                {
                    return File.Exists(targetPath);
                }

                var generated = activeView.GenerateImage(style, size.Width, size.Height);
                if (TrySaveNavisImageToJpeg(generated, targetPath))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to render viewpoint image for {viewpoint.DisplayName}: {ex.Message}");
            }

            return false;
        }

        private static bool TryApplyViewpoint(Document document, SavedViewpoint viewpoint)
        {
            try
            {
                var applied = false;
                var savedViewpoints = document.SavedViewpoints;
                if (savedViewpoints != null)
                {
                    var currentProp = savedViewpoints.GetType().GetProperty("CurrentSavedViewpoint");
                    if (currentProp != null && currentProp.CanWrite)
                    {
                        currentProp.SetValue(savedViewpoints, viewpoint);
                        applied = true;
                    }
                }

                var applyMethod = viewpoint.GetType().GetMethod("ApplyToDocument", new[] { typeof(Document) });
                if (applyMethod != null)
                {
                    applyMethod.Invoke(viewpoint, new object[] { document });
                    applied = true;
                }

                var viewpointProperty = viewpoint.GetType().GetProperty("Viewpoint");
                var navisViewpoint = viewpointProperty?.GetValue(viewpoint);
                if (navisViewpoint != null)
                {
                    var documentType = document.GetType();
                    var currentViewpointProperty = documentType.GetProperty("CurrentViewpoint");
                    if (currentViewpointProperty != null && currentViewpointProperty.CanWrite)
                    {
                        currentViewpointProperty.SetValue(document, navisViewpoint);
                        applied = true;
                    }
                }

                return applied;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to apply viewpoint {viewpoint.DisplayName}: {ex.Message}");
            }

            return false;
        }

        private static Bitmap? TryGenerateThumbnail(SavedViewpoint viewpoint, Size size)
        {
            try
            {
                var type = viewpoint.GetType();

                var sizeMethod = type.GetMethod("GenerateThumbnail", new[] { typeof(Size) });
                var sizedResult = sizeMethod?.Invoke(viewpoint, new object[] { size });
                var sizedBitmap = TryConvertToBitmap(sizedResult);
                if (sizedBitmap != null)
                {
                    return sizedBitmap;
                }

                var noArgMethod = type.GetMethod("GenerateThumbnail", Type.EmptyTypes);
                var noArgResult = noArgMethod?.Invoke(viewpoint, Array.Empty<object>());
                var noArgBitmap = TryConvertToBitmap(noArgResult);
                if (noArgBitmap != null)
                {
                    return noArgBitmap;
                }

                var property = type.GetProperty("Thumbnail");
                var propertyResult = property?.GetValue(viewpoint);
                var propertyBitmap = TryConvertToBitmap(propertyResult);
                if (propertyBitmap != null)
                {
                    return propertyBitmap;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Unable to create thumbnail for viewpoint {viewpoint.DisplayName}: {ex.Message}");
            }

            return null;
        }

        private static void SaveBitmapToJpeg(Bitmap bitmap, string targetPath)
        {
            bitmap.Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private static bool TrySaveNavisImageToJpeg(Autodesk.Navisworks.Api.Image? navisImage, string targetPath)
        {
            if (navisImage == null)
            {
                return false;
            }

            try
            {
                using (navisImage)
                {
                    try
                    {
                        navisImage.Save(targetPath, ImageFileType.Jpeg);
                        return File.Exists(targetPath);
                    }
                    catch (MissingMethodException)
                    {
                        using var bitmap = TryConvertNavisImageToBitmap(navisImage);
                        if (bitmap == null)
                        {
                            return false;
                        }

                        SaveBitmapToJpeg(bitmap, targetPath);
                        return File.Exists(targetPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to save Navisworks image to {targetPath}: {ex.Message}");
            }

            return false;
        }

        private static Bitmap? TryConvertToBitmap(object? candidate)
        {
            switch (candidate)
            {
                case null:
                    return null;
                case Bitmap bitmap:
                    return bitmap;
                case Autodesk.Navisworks.Api.Image navisImage:
                    return TryConvertNavisImageToBitmap(navisImage);
            }

            var candidateType = candidate.GetType();
            var fullName = candidateType.FullName ?? string.Empty;
            if (string.Equals(fullName, "Autodesk.Navisworks.Api.Interop.LcOpSavedItemCommentImage", StringComparison.Ordinal))
            {
                var tempPath = Path.GetTempFileName();
                try
                {
                    var saveMethod = candidateType.GetMethod("Save", new[] { typeof(string), typeof(ImageFileType) });
                    if (saveMethod != null)
                    {
                        saveMethod.Invoke(candidate, new object[] { tempPath, ImageFileType.Jpeg });
                        if (File.Exists(tempPath))
                        {
                            return new Bitmap(tempPath);
                        }
                    }

                    var toBitmapMethod = candidateType.GetMethod("ToBitmap", Type.EmptyTypes);
                    if (toBitmapMethod?.Invoke(candidate, Array.Empty<object>()) is Bitmap reflectedBitmap)
                    {
                        return reflectedBitmap;
                    }
                }
                finally
                {
                    try
                    {
                        if (File.Exists(tempPath))
                        {
                            File.Delete(tempPath);
                        }
                    }
                    catch
                    {
                        // ignore cleanup failures
                    }
                }
            }

            return null;
        }

        private static Bitmap? TryConvertNavisImageToBitmap(Autodesk.Navisworks.Api.Image navisImage)
        {
            try
            {
                using var memory = new MemoryStream();
                var imageType = navisImage.GetType();
                var saveToStream = imageType.GetMethod("Save", new[] { typeof(Stream), typeof(ImageFileType) });
                if (saveToStream != null)
                {
                    saveToStream.Invoke(navisImage, new object[] { memory, ImageFileType.Jpeg });
                    memory.Position = 0;
                    return new Bitmap(memory);
                }

                var toBitmap = imageType.GetMethod("ToBitmap", Type.EmptyTypes);
                if (toBitmap?.Invoke(navisImage, Array.Empty<object>()) is Bitmap bitmap)
                {
                    return bitmap;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to convert Navisworks image to bitmap: {ex.Message}");
            }

            return null;
        }
    }
}
