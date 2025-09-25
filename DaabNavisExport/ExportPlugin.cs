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

                var normalizedPath = imagePath!.Trim();
                if (normalizedPath.Length == 0)
                {
                    continue;
                }

                if (!imageAssignments.ContainsKey(guid))
                {
                    imageAssignments.Add(guid, normalizedPath);
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

                if (TryRenderViewpointImage(context.Document, viewpoint, targetPath, new Size(800, 450)))
                {
                    continue;
                }

                if (TryGenerateThumbnail(viewpoint, targetPath, new Size(800, 450)))
                {
                    if (!File.Exists(targetPath))
                    {
                        Debug.WriteLine($"Thumbnail generation reported success but file not found at {targetPath}.");
                    }

                    continue;
                }

                Debug.WriteLine($"No renderer succeeded for viewpoint {viewpoint.DisplayName} (GUID={viewpoint.Guid}).");
            }
        }

        private static bool TryRenderViewpointImage(Document document, SavedViewpoint viewpoint, string targetPath, Size size)
        {
            try
            {
                TryApplyViewpoint(document, viewpoint);

                var activeViewProperty = document.GetType().GetProperty("ActiveView");
                var activeView = activeViewProperty?.GetValue(document);
                if (activeView == null)
                {
                    return false;
                }

                var viewType = activeView.GetType();

                if (TryInvokeViewToFile(activeView, viewType, "RenderToImage", targetPath, size))
                {
                    return true;
                }

                if (TryInvokeViewToFile(activeView, viewType, "SaveToImage", targetPath, size))
                {
                    return true;
                }

                if (TryInvokeViewWithStyle(activeView, viewType, "RenderToImage", targetPath, size))
                {
                    return true;
                }

                if (TryInvokeViewWithStyle(activeView, viewType, "SaveToImage", targetPath, size))
                {
                    return true;
                }

                if (TryGenerateImage(activeView, viewType, targetPath, size))
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

        private static bool TryGenerateThumbnail(SavedViewpoint viewpoint, string targetPath, Size size)
        {
            try
            {
                var type = viewpoint.GetType();

                var sizeMethod = type.GetMethod("GenerateThumbnail", new[] { typeof(Size) });
                if (sizeMethod != null)
                {
                    var result = sizeMethod.Invoke(viewpoint, new object[] { size });
                    if (TrySaveImageToPath(result, targetPath))
                    {
                        return true;
                    }
                }

                var noArgMethod = type.GetMethod("GenerateThumbnail", Type.EmptyTypes);
                if (noArgMethod != null)
                {
                    var result = noArgMethod.Invoke(viewpoint, Array.Empty<object>());
                    if (TrySaveImageToPath(result, targetPath))
                    {
                        return true;
                    }
                }

                var property = type.GetProperty("Thumbnail");
                if (property != null)
                {
                    var value = property.GetValue(viewpoint);
                    if (TrySaveImageToPath(value, targetPath))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unable to create thumbnail for viewpoint {viewpoint.DisplayName}: {ex.Message}");
            }

            return false;
        }

        private static bool TryInvokeViewToFile(object activeView, Type viewType, string methodName, string targetPath, Size size)
        {
            var method = viewType.GetMethod(methodName, new[] { typeof(string), typeof(int), typeof(int) });
            if (method == null)
            {
                return false;
            }

            method.Invoke(activeView, new object[] { targetPath, size.Width, size.Height });
            return File.Exists(targetPath);
        }

        private static bool TryInvokeViewWithStyle(object activeView, Type viewType, string methodName, string targetPath, Size size)
        {
            foreach (var method in viewType.GetMethods().Where(m => m.Name == methodName))
            {
                var parameters = method.GetParameters();
                if (parameters.Length != 4)
                {
                    continue;
                }

                if (parameters[0].ParameterType != typeof(string) ||
                    parameters[2].ParameterType != typeof(int) ||
                    parameters[3].ParameterType != typeof(int))
                {
                    continue;
                }

                var styleValue = ResolveEnumValue(parameters[1].ParameterType, new[] { "Raster", "Standard", "Smooth", "HighQuality" });
                if (styleValue == null)
                {
                    if (parameters[1].ParameterType.IsValueType)
                    {
                        styleValue = Activator.CreateInstance(parameters[1].ParameterType);
                    }
                    else
                    {
                        continue;
                    }
                }

                method.Invoke(activeView, new[] { targetPath, styleValue, size.Width, size.Height });
                if (File.Exists(targetPath))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGenerateImage(object activeView, Type viewType, string targetPath, Size size)
        {
            foreach (var method in viewType.GetMethods().Where(m => m.Name == "GenerateImage"))
            {
                var parameters = method.GetParameters();
                object? result = null;

                if (parameters.Length == 2 &&
                    parameters[0].ParameterType == typeof(int) &&
                    parameters[1].ParameterType == typeof(int))
                {
                    result = method.Invoke(activeView, new object[] { size.Width, size.Height });
                }
                else if (parameters.Length == 3 &&
                         parameters[1].ParameterType == typeof(int) &&
                         parameters[2].ParameterType == typeof(int))
                {
                    var styleValue = ResolveEnumValue(parameters[0].ParameterType, new[] { "Raster", "Standard", "Smooth", "HighQuality" });
                    if (styleValue == null)
                    {
                        if (parameters[0].ParameterType.IsValueType)
                        {
                            styleValue = Activator.CreateInstance(parameters[0].ParameterType);
                        }
                        else
                        {
                            continue;
                        }
                    }

                    result = method.Invoke(activeView, new[] { styleValue, size.Width, size.Height });
                }

                if (TrySaveImageToPath(result, targetPath))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TrySaveImageToPath(object? imageObject, string targetPath)
        {
            if (imageObject == null)
            {
                return false;
            }

            if (imageObject is Bitmap bitmap)
            {
                using (bitmap)
                {
                    bitmap.Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                return File.Exists(targetPath);
            }

            var disposable = imageObject as IDisposable;
            try
            {
                if (TrySaveViaReflection(imageObject, targetPath))
                {
                    return true;
                }

                var toBitmapMethod = imageObject.GetType().GetMethod("ToBitmap", Type.EmptyTypes);
                if (toBitmapMethod?.Invoke(imageObject, Array.Empty<object>()) is Bitmap converted)
                {
                    using (converted)
                    {
                        converted.Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }

                    return File.Exists(targetPath);
                }
            }
            finally
            {
                disposable?.Dispose();
            }

            return false;
        }

        private static bool TrySaveViaReflection(object imageObject, string targetPath)
        {
            var type = imageObject.GetType();

            var saveString = type.GetMethod("Save", new[] { typeof(string) });
            if (saveString != null)
            {
                saveString.Invoke(imageObject, new object[] { targetPath });
                if (File.Exists(targetPath))
                {
                    return true;
                }
            }

            foreach (var method in type.GetMethods().Where(m => m.Name == "Save"))
            {
                var parameters = method.GetParameters();
                if (parameters.Length != 2 || parameters[0].ParameterType != typeof(string))
                {
                    continue;
                }

                var formatValue = ResolveEnumValue(parameters[1].ParameterType, new[] { "Jpeg", "JPEG", "Jpg" });
                if (formatValue == null)
                {
                    if (parameters[1].ParameterType.IsValueType)
                    {
                        formatValue = Activator.CreateInstance(parameters[1].ParameterType);
                    }
                    else
                    {
                        continue;
                    }
                }

                method.Invoke(imageObject, new[] { targetPath, formatValue });
                if (File.Exists(targetPath))
                {
                    return true;
                }
            }

            var writeToFile = type.GetMethod("WriteToFile", new[] { typeof(string) });
            if (writeToFile != null)
            {
                writeToFile.Invoke(imageObject, new object[] { targetPath });
                return File.Exists(targetPath);
            }

            return false;
        }

        private static object? ResolveEnumValue(Type enumType, IReadOnlyList<string> preferredNames)
        {
            if (!enumType.IsEnum)
            {
                return null;
            }

            var names = Enum.GetNames(enumType);
            foreach (var preferred in preferredNames)
            {
                var match = names.FirstOrDefault(n => string.Equals(n, preferred, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    return Enum.Parse(enumType, match);
                }
            }

            return names.Length > 0 ? Enum.Parse(enumType, names[0]) : null;
        }
        }
}
