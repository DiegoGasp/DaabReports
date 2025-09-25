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
            ExportContext? exportContext = null;
            ProgressReporter? progress = null;
            var completedSteps = 0;

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
                exportContext = BuildExportContext(document, outputDirectory);

                exportContext.Log.Info($"Using output directory: {outputDirectory}");
                exportContext.Log.Info($"Document filename: {document.FileName ?? "<unsaved>"}");

                progress = new ProgressReporter(4);
                progress.Report(completedSteps, "Writing viewpoints to XML...");

                ExportViewpointsToXml(document, exportContext);
                completedSteps++;
                exportContext.Log.Info($"Captured {exportContext.ViewSequence.Count} viewpoints.");

                var totalSteps = Math.Max(exportContext.ViewSequence.Count + 4, 4);
                progress.UpdateMaximum(totalSteps);
                progress.Report(completedSteps, "Parsing exported XML...");

                var parser = new NavisworksXmlParser();
                var parseResult = parser.Process(exportContext.XmlPath);
                exportContext.Log.Info($"Parsed {parseResult.Rows.Count} CSV rows.");

                completedSteps++;
                progress.Report(completedSteps, "Writing DB outputs...");
                parser.WriteOutputs(parseResult, exportContext.DbDirectory);
                exportContext.Log.Info("Database outputs written successfully.");

                completedSteps++;
                progress.Report(completedSteps, "Rendering viewpoint images...");
                completedSteps = ExportViewpointImages(exportContext, parseResult.Rows, progress, completedSteps);

                progress.Report(totalSteps, "Export complete.");
                exportContext.Log.Info("Export finished successfully.");

                MessageBox.Show(
                    $"Export complete.\nProject folder: {exportContext.ProjectDirectory}\nXML: {exportContext.XmlPath}\nCSV: {Path.Combine(exportContext.DbDirectory, NavisworksXmlParser.CsvFileName)}\nLog: {exportContext.Log.FilePath}",
                    "Daab Navis Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return 0;
            }
            catch (Exception ex)
            {
                exportContext?.Log.Error($"Export failed: {ex}");
                MessageBox.Show($"Export failed: {ex.Message}", "Daab Navis Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
            finally
            {
                progress?.Dispose();
                exportContext?.Log.Complete();
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
            var logFile = Path.Combine(projectDirectory, $"export_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            var log = new ExportLog(logFile);

            log.Info($"Project directory: {projectDirectory}");
            log.Info($"DB directory: {dbDirectory}");
            log.Info($"Images directory: {imagesDirectory}");
            log.Info($"XML output: {xmlFile}");

            return new ExportContext(document, outputDirectory, projectDirectory, dbDirectory, imagesDirectory, xmlFile, log);
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
                return explicitPath!;
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

        private static int ExportViewpointImages(
            ExportContext context,
            IEnumerable<IReadOnlyList<string?>> rows,
            ProgressReporter progress,
            int completedSteps)
        {
            if (context.ViewSequence.Count == 0)
            {
                context.Log.Warn("No viewpoints discovered while attempting to export images.");
                return completedSteps;
            }

            Directory.CreateDirectory(context.ImagesDirectory);

            var log = context.Log;
            var imageAssignments = new Dictionary<Guid, string>();
            var guidIndex = -1;
            var imageIndex = -1;
            var headerProcessed = false;

            foreach (var row in rows)
            {
                if (!headerProcessed)
                {
                    for (var i = 0; i < row.Count; i++)
                    {
                        var header = row[i];
                        if (string.Equals(header, "GUID", StringComparison.OrdinalIgnoreCase))
                        {
                            guidIndex = i;
                        }
                        else if (string.Equals(header, "Image", StringComparison.OrdinalIgnoreCase))
                        {
                            imageIndex = i;
                        }
                    }

                    headerProcessed = true;
                    continue;
                }

                if (guidIndex == -1 || imageIndex == -1)
                {
                    continue;
                }

                var guidValue = row.Count > guidIndex ? row[guidIndex] : null;
                var imageValue = row.Count > imageIndex ? row[imageIndex] : null;

                if (Guid.TryParse(guidValue, out var guid) && !string.IsNullOrEmpty(imageValue))
                {
                    var normalizedPath = imageValue.Replace('/', Path.DirectorySeparatorChar);
                    log.Info($"Row mapped viewpoint {guid} to image '{normalizedPath}'.");

                    if (!imageAssignments.ContainsKey(guid))
                    {
                        imageAssignments.Add(guid, normalizedPath);
                    }
                    else
                    {
                        log.Warn($"Duplicate image mapping for viewpoint {guid} ignored (existing: '{imageAssignments[guid]}', incoming: '{normalizedPath}').");
                    }
                }
            }

            var totalViewpoints = context.ViewSequence.Count;
            log.Info($"Preparing to render {totalViewpoints} viewpoints.");

            var index = 0;
            foreach (var viewpoint in context.ViewSequence)
            {
                if (!imageAssignments.TryGetValue(viewpoint.Guid, out var imageFile))
                {
                    log.Warn($"No image assignment found for viewpoint {viewpoint.Guid}; skipping.");
                    continue;
                }

                var targetPath = Path.Combine(context.ImagesDirectory, imageFile);
                var targetDirectory = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                var displayName = viewpoint.DisplayName ?? viewpoint.Guid.ToString();
                var progressMessage = $"Rendering {displayName} ({index + 1}/{totalViewpoints})";
                progress.Report(completedSteps, progressMessage + "...");
                log.Info($"Rendering viewpoint '{displayName}' to '{targetPath}'.");

                var rendered = TryRenderViewpointImage(context.Document, viewpoint, targetPath, new Size(800, 450), log);
                if (!rendered)
                {
                    rendered = TryGenerateThumbnail(viewpoint, targetPath, new Size(800, 450), log);
                    if (rendered && !File.Exists(targetPath))
                    {
                        log.Warn($"Thumbnail generation reported success but file not found at {targetPath}.");
                        rendered = false;
                    }
                }

                completedSteps++;
                index++;

                if (!rendered)
                {
                    var failureMessage = $"No renderer succeeded for viewpoint {displayName} (GUID={viewpoint.Guid}).";
                    log.Warn(failureMessage);
                    progress.Report(completedSteps, failureMessage);
                }
                else
                {
                    var successMessage = $"Saved viewpoint '{displayName}' to '{targetPath}'.";
                    log.Info(successMessage);
                    progress.Report(completedSteps, successMessage);
                }
            }

            return completedSteps;
        }

        private static bool TryRenderViewpointImage(Document document, SavedViewpoint viewpoint, string targetPath, Size size, ExportLog log)
        {
            try
            {
                if (!TryApplyViewpoint(document, viewpoint, log))
                {
                    log.Warn($"Viewpoint {viewpoint.Guid} could not be fully applied; attempting render with current view state.");
                }

                var activeViewProperty = document.GetType().GetProperty("ActiveView");
                var activeView = activeViewProperty?.GetValue(document);
                if (activeView == null)
                {
                    log.Warn("ActiveView property not available on document.");
                    return false;
                }

                var viewType = activeView.GetType();

                log.Info("Attempting RenderToImage(string, int, int) via reflection.");
                if (TryInvokeViewToFile(activeView, viewType, "RenderToImage", targetPath, size, log))
                {
                    log.Info("RenderToImage(string, int, int) succeeded.");
                    return true;
                }

                log.Info("Attempting SaveToImage(string, int, int) via reflection.");
                if (TryInvokeViewToFile(activeView, viewType, "SaveToImage", targetPath, size, log))
                {
                    log.Info("SaveToImage(string, int, int) succeeded.");
                    return true;
                }

                log.Info("Attempting RenderToImage overloads with style parameter.");
                if (TryInvokeViewWithStyle(activeView, viewType, "RenderToImage", targetPath, size, log))
                {
                    log.Info("RenderToImage(style, string, int, int) succeeded.");
                    return true;
                }

                log.Info("Attempting SaveToImage overloads with style parameter.");
                if (TryInvokeViewWithStyle(activeView, viewType, "SaveToImage", targetPath, size, log))
                {
                    log.Info("SaveToImage(style, string, int, int) succeeded.");
                    return true;
                }

                log.Info("Attempting GenerateImage methods.");
                if (TryGenerateImage(activeView, viewType, targetPath, size, log))
                {
                    log.Info("GenerateImage methods produced an image.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                log.Error($"Failed to render viewpoint image for {viewpoint.DisplayName}: {ex}");
            }

            return false;
        }

        private static bool TryApplyViewpoint(Document document, SavedViewpoint viewpoint, ExportLog log)
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
                        log.Info($"Set CurrentSavedViewpoint to {viewpoint.DisplayName ?? viewpoint.Guid.ToString()}.");
                    }
                }

                var applyMethod = viewpoint.GetType().GetMethod("ApplyToDocument", new[] { typeof(Document) });
                if (applyMethod != null)
                {
                    applyMethod.Invoke(viewpoint, new object[] { document });
                    applied = true;
                    log.Info("Invoked ApplyToDocument on SavedViewpoint.");
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
                        log.Info("Updated Document.CurrentViewpoint via reflection.");
                    }
                }

                return applied;
            }
            catch (Exception ex)
            {
                log.Error($"Failed to apply viewpoint {viewpoint.DisplayName}: {ex}");
            }

            return false;
        }

        private static bool TryGenerateThumbnail(SavedViewpoint viewpoint, string targetPath, Size size, ExportLog log)
        {
            try
            {
                var type = viewpoint.GetType();

                var sizeMethod = type.GetMethod("GenerateThumbnail", new[] { typeof(Size) });
                if (sizeMethod != null)
                {
                    var result = sizeMethod.Invoke(viewpoint, new object[] { size });
                    log.Info("Invoked GenerateThumbnail(Size).");
                    if (TrySaveImageToPath(result, targetPath, log))
                    {
                        return true;
                    }
                }

                var noArgMethod = type.GetMethod("GenerateThumbnail", Type.EmptyTypes);
                if (noArgMethod != null)
                {
                    var result = noArgMethod.Invoke(viewpoint, Array.Empty<object>());
                    log.Info("Invoked GenerateThumbnail().");
                    if (TrySaveImageToPath(result, targetPath, log))
                    {
                        return true;
                    }
                }

                var property = type.GetProperty("Thumbnail");
                if (property != null)
                {
                    var value = property.GetValue(viewpoint);
                    log.Info("Attempting to save Thumbnail property value.");
                    if (TrySaveImageToPath(value, targetPath, log))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error($"Unable to create thumbnail for viewpoint {viewpoint.DisplayName}: {ex}");
            }

            return false;
        }

        private static bool TryInvokeViewToFile(object activeView, Type viewType, string methodName, string targetPath, Size size, ExportLog log)
        {
            var method = viewType.GetMethod(methodName, new[] { typeof(string), typeof(int), typeof(int) });
            if (method == null)
            {
                log.Info($"Method {methodName}(string, int, int) not found on {viewType.FullName}.");
                return false;
            }

            method.Invoke(activeView, new object[] { targetPath, size.Width, size.Height });
            var created = File.Exists(targetPath);
            if (!created)
            {
                log.Warn($"{methodName}(string, int, int) did not create a file at {targetPath}.");
            }

            return created;
        }

        private static bool TryInvokeViewWithStyle(object activeView, Type viewType, string methodName, string targetPath, Size size, ExportLog log)
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

                log.Info($"Invoking {method.Name}(string, {parameters[1].ParameterType.Name}, int, int) with style {styleValue}.");
                method.Invoke(activeView, new[] { targetPath, styleValue, size.Width, size.Height });
                if (File.Exists(targetPath))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGenerateImage(object activeView, Type viewType, string targetPath, Size size, ExportLog log)
        {
            foreach (var method in viewType.GetMethods().Where(m => m.Name == "GenerateImage"))
            {
                var parameters = method.GetParameters();
                object? result = null;

                if (parameters.Length == 2 &&
                    parameters[0].ParameterType == typeof(int) &&
                    parameters[1].ParameterType == typeof(int))
                {
                    log.Info("Invoking GenerateImage(int width, int height).");
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

                    log.Info($"Invoking GenerateImage({parameters[0].ParameterType.Name} style, int width, int height) with style {styleValue}.");
                    result = method.Invoke(activeView, new[] { styleValue, size.Width, size.Height });
                }

                if (TrySaveImageToPath(result, targetPath, log))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TrySaveImageToPath(object? imageObject, string targetPath, ExportLog log)
        {
            if (imageObject == null)
            {
                log.Warn("Image object was null; nothing to save.");
                return false;
            }

            if (imageObject is Bitmap bitmap)
            {
                using (bitmap)
                {
                    bitmap.Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                log.Info($"Saved System.Drawing.Bitmap to {targetPath}.");
                return File.Exists(targetPath);
            }

            var disposable = imageObject as IDisposable;
            try
            {
                if (TrySaveViaReflection(imageObject, targetPath, log))
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

                    log.Info($"Converted {imageObject.GetType().FullName} to Bitmap and saved to {targetPath}.");
                    return File.Exists(targetPath);
                }
            }
            finally
            {
                disposable?.Dispose();
            }

            return false;
        }

        private static bool TrySaveViaReflection(object imageObject, string targetPath, ExportLog log)
        {
            var type = imageObject.GetType();

            var saveString = type.GetMethod("Save", new[] { typeof(string) });
            if (saveString != null)
            {
                saveString.Invoke(imageObject, new object[] { targetPath });
                if (File.Exists(targetPath))
                {
                    log.Info($"Saved {type.FullName} via Save(string) to {targetPath}.");
                    return true;
                }
                log.Warn($"Save(string) on {type.FullName} did not create {targetPath}.");
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
                    log.Info($"Saved {type.FullName} via Save(string, {parameters[1].ParameterType.Name}) to {targetPath}.");
                    return true;
                }
                log.Warn($"Save(string, {parameters[1].ParameterType.Name}) on {type.FullName} did not create {targetPath}.");
            }

            var writeToFile = type.GetMethod("WriteToFile", new[] { typeof(string) });
            if (writeToFile != null)
            {
                writeToFile.Invoke(imageObject, new object[] { targetPath });
                var created = File.Exists(targetPath);
                if (created)
                {
                    log.Info($"Saved {type.FullName} via WriteToFile(string) to {targetPath}.");
                }
                else
                {
                    log.Warn($"WriteToFile(string) on {type.FullName} did not create {targetPath}.");
                }

                return created;
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
