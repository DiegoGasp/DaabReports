using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Data;
using Autodesk.Navisworks.Api.Plugins;
using CloudinaryDotNet;
using CloudinaryDotNet.Actions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NavisApplication = Autodesk.Navisworks.Api.Application;

namespace DaabNavisExport
{
    // Ribbon layout
    [Plugin("DaabRibbon", "DAAB")]
    [RibbonLayout("DaabRibbon.xaml")]
    [Command("RunExport", DisplayName = "Export Report", Icon = "MascotIcon.png", LargeIcon = "MascotIcon.png", ToolTip = "Export viewpoints and comments to report")]
    [Command("OpenSettings", DisplayName = "Settings", Icon = "MascotIcon.png", LargeIcon = "MascotIcon.png", ToolTip = "Configure export settings")]
    [Command("CreateViewpoint", DisplayName = "Create Viewpoint", Icon = "MascotIcon.png", LargeIcon = "MascotIcon.png", ToolTip = "Create and organize viewpoint")]
    public sealed class DaabRibbon : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string commandId, params string[] parameters)
        {
            switch (commandId)
            {
                case "RunExport":
                    NavisApplication.Plugins.ExecuteAddInPlugin("DaabNavisExport.DAAB");
                    break;
                case "OpenSettings":
                    ShowSettingsDialog();
                    break;
                case "CreateViewpoint":
                    NavisApplication.Plugins.ExecuteAddInPlugin("DaabViewpointCreator.DAAB");
                    break;
            }
            return 0;
        }

        private void ShowSettingsDialog()
        {
            try
            {
                // Get Navisworks main window handle
                var mainWindow = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                if (mainWindow != IntPtr.Zero)
                {
                    MessageBox.Show("No active document open.", "Daab Navis Export",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }
                else
                {
                    dialog.ShowDialog();
                }
            }
        }
    }

    // Settings dialog
    public class SettingsForm : Form
    {
        private NumericUpDown widthInput = null!;
        private NumericUpDown heightInput = null!;
        private CheckBox cloudinaryCheckbox = null!;
        private ListBox subfolderListBox = null!;
        private TextBox newSubfolderText = null!;
        private Button addSubfolderButton = null!;
        private Button removeSubfolderButton = null!;
        private Button saveButton = null!;
        private Button cancelButton = null!;

        public SettingsForm()
        {
            InitializeComponents();
            LoadSettings();
        }

        private void InitializeComponents()
        {
            Text = "Daab Export Settings";
            Width = 450;
            Height = 450;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            var labelWidth = new Label { Text = "Image Width:", Left = 20, Top = 20, Width = 120 };
            widthInput = new NumericUpDown { Left = 150, Top = 20, Width = 250, Minimum = 640, Maximum = 3840, Increment = 160 };

            var labelHeight = new Label { Text = "Image Height:", Left = 20, Top = 60, Width = 120 };
            heightInput = new NumericUpDown { Left = 150, Top = 60, Width = 250, Minimum = 480, Maximum = 2160, Increment = 120 };

            cloudinaryCheckbox = new CheckBox { Text = "Enable Cloudinary Upload", Left = 20, Top = 100, Width = 300 };

            var labelSubfolders = new Label { Text = "Subfolder Options:", Left = 20, Top = 140, Width = 380 };
            subfolderListBox = new ListBox { Left = 20, Top = 165, Width = 380, Height = 120 };

            var projectFolderName = ResolveProjectFolderName(document);
            var projectDirectory = Path.Combine(outputDirectory, projectFolderName);
            var dbDirectory = Path.Combine(projectDirectory, DbFolderName);
            var imagesDirectory = Path.Combine(projectDirectory, ImagesFolderName);

            Directory.CreateDirectory(projectDirectory);
            Directory.CreateDirectory(dbDirectory);
            Directory.CreateDirectory(imagesDirectory);

            var xmlFile = Path.Combine(dbDirectory, "DB.xml");

            cancelButton = new Button { Text = "Cancel", Left = 320, Top = 350, Width = 80 };
            cancelButton.Click += CancelButton_Click;

            Controls.AddRange(new Control[] {
                labelWidth, widthInput, labelHeight, heightInput, cloudinaryCheckbox,
                labelSubfolders, subfolderListBox, newSubfolderText, addSubfolderButton,
                removeSubfolderButton, saveButton, cancelButton
            });

            AcceptButton = saveButton;
            CancelButton = cancelButton;
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

        private static string ResolveOutputDirectory(IReadOnlyList<string> parameters)
        {
            var explicitPath = parameters?.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
            if (!string.IsNullOrEmpty(explicitPath))
            {
                return explicitPath!;
            }

        private void LoadSettings()
        {
            widthInput.Value = ExportSettings.Instance.ImageWidth;
            heightInput.Value = ExportSettings.Instance.ImageHeight;
            cloudinaryCheckbox.Checked = ExportSettings.Instance.EnableCloudinary;

            subfolderListBox.Items.Clear();
            foreach (var option in ExportSettings.Instance.SubfolderOptions)
                subfolderListBox.Items.Add(option);
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
            catch { }
            }

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
                            case "EnableCloudinary":
                                if (bool.TryParse(parts[1], out bool c)) settings.EnableCloudinary = c;
                                break;
                            case "SubfolderOptions":
                                if (!string.IsNullOrWhiteSpace(parts[1]))
                                    settings.SubfolderOptions = parts[1].Split('|').ToList();
                                break;
            }
        }
            }
        }
            catch { }
            return settings;
        }
    }

    // Viewpoint creator plugin
    [Plugin("DaabViewpointCreator", "DAAB", DisplayName = "Create Viewpoint", ToolTip = "Create organized viewpoint")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ViewpointCreatorPlugin : AddInPlugin
        {
        public override int Execute(params string[] parameters)
        {
            try
            {
                var commentsProperty = viewpoint.GetType().GetProperty("Comments");
                if (commentsProperty == null)
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

        private static void WriteCreatedDate(XmlWriter writer, object? created)
        {
            if (created is not DateTime createdDate)
            {
                return;
            }

            if (createdDate.Year < 1900)
            {
                return;

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
                }

                var rendered = TryRenderViewpointImage(context.Document, viewpoint, targetPath, new Size(800, 450));
                if (!rendered)
                {
                    rendered = TryGenerateThumbnail(viewpoint, targetPath, new Size(800, 450));
                    if (rendered && !File.Exists(targetPath))
                    {
                        Debug.WriteLine($"Thumbnail generation reported success but file not found at {targetPath}.");
                        rendered = false;
                    }
                }

                if (!rendered)
                {
                    Debug.WriteLine($"No renderer succeeded for viewpoint {viewpoint.DisplayName} (GUID={viewpoint.Guid}).");
                }
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

                var safeProjectName = ToSafeFileName(_projectName);
                var destFileName = $"{safeProjectName}_Report.pbix";
                var dest = Path.Combine(exportContext.ProjectDirectory, destFileName);

                Directory.CreateDirectory(exportContext.ProjectDirectory);
                File.Copy(src, dest, overwrite: true);

                if (File.Exists(dest))
                    Log("Power BI template copied to: " + dest);
                else
                    Log("PBIX copy failed - file not found after copy.");
                }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to apply viewpoint {viewpoint.DisplayName}: {ex.Message}");
            }
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
            return "Navisworks_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                }

                var noArgMethod = type.GetMethod("GenerateThumbnail", Type.EmptyTypes);
                if (noArgMethod != null)
                {
                    var result = noArgMethod.Invoke(viewpoint, Array.Empty<object>());
                    if (TrySaveImageToPath(result, targetPath))
                    {
                        return true;
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
                    }

                method.Invoke(activeView, new[] { targetPath, styleValue, size.Width, size.Height });
                if (File.Exists(targetPath))
                {
                    return true;
                }
            }
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
                }

                if (TrySaveImageToPath(result, targetPath))
                {
                    return true;
                }
            }

            return false;
        }

            if (imageObject is Bitmap bitmap)
            {
                using (bitmap)
                {
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
                }
            }
            finally
            {
                disposable?.Dispose();
            }
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


            var writeToFile = type.GetMethod("WriteToFile", new[] { typeof(string) });
            if (writeToFile != null)
            {
                writeToFile.Invoke(imageObject, new object[] { targetPath });
                return File.Exists(targetPath);
            }

            return false;
        }
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
