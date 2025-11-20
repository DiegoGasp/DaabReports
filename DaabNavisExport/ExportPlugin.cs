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
    internal sealed class HwndWrapper : IWin32Window
    {
        public IntPtr Handle { get; }
        public HwndWrapper(IntPtr handle) => Handle = handle;
    }

    // Ribbon layout
    [Plugin("DaabRibbon", "DAAB")]
    [RibbonLayout("DaabRibbon.xaml")]
    [RibbonTab("ID_DaabTab_1", DisplayName = "Daab Reports")] // matches RibbonTab Id
    [Command("RunExport", DisplayName = "Export Report", Icon = "ReportIco.png", LargeIcon = "ReportIco.png")]
    [Command("OpenSettings", DisplayName = "Settings", Icon = "SettingsIco.png", LargeIcon = "SettingsIco.png")]
    [Command("CreateViewpoint", DisplayName = "Create Viewpoint", Icon = "CameraIco.png", LargeIcon = "CameraIco.png")]
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
            using (var dialog = new SettingsForm())
            {
                var mainWindow = Process.GetCurrentProcess().MainWindowHandle;
                if (mainWindow != IntPtr.Zero)
                {
                    dialog.ShowDialog(new HwndWrapper(mainWindow));
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

            newSubfolderText = new TextBox { Left = 20, Top = 295, Width = 280 };
            addSubfolderButton = new Button { Text = "Add", Left = 310, Top = 293, Width = 45 };
            addSubfolderButton.Click += AddSubfolderButton_Click;

            removeSubfolderButton = new Button { Text = "Remove", Left = 360, Top = 293, Width = 40 };
            removeSubfolderButton.Click += RemoveSubfolderButton_Click;

            saveButton = new Button
            {
                Text = "Save",
                Left = 230,
                Top = 350,
                Width = 80,
                DialogResult = DialogResult.OK
            };
            saveButton.Click += SaveButton_Click;

            cancelButton = new Button
            {
                Text = "Cancel",
                Left = 320,
                Top = 350,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };
            cancelButton.Click += CancelButton_Click;

            Controls.AddRange(new Control[] {
                labelWidth, widthInput, labelHeight, heightInput, cloudinaryCheckbox,
                labelSubfolders, subfolderListBox, newSubfolderText, addSubfolderButton,
                removeSubfolderButton, saveButton, cancelButton
            });

            AcceptButton = saveButton;
            CancelButton = cancelButton;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void AddSubfolderButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(newSubfolderText.Text))
            {
                subfolderListBox.Items.Add(newSubfolderText.Text);
                newSubfolderText.Clear();
            }
        }

        private void RemoveSubfolderButton_Click(object sender, EventArgs e)
        {
            if (subfolderListBox.SelectedItem != null)
                subfolderListBox.Items.Remove(subfolderListBox.SelectedItem);
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

        private void SaveButton_Click(object sender, EventArgs e)
        {
            ExportSettings.Instance.ImageWidth = (int)widthInput.Value;
            ExportSettings.Instance.ImageHeight = (int)heightInput.Value;
            ExportSettings.Instance.EnableCloudinary = cloudinaryCheckbox.Checked;

            ExportSettings.Instance.SubfolderOptions.Clear();
            foreach (var item in subfolderListBox.Items)
                ExportSettings.Instance.SubfolderOptions.Add(item.ToString() ?? "");

            ExportSettings.Instance.Save();
            MessageBox.Show("Settings saved successfully!", "Daab Export",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // Settings storage
    public class ExportSettings
    {
        private static ExportSettings? _instance;
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DaabNavisExport",
            "settings.txt");

        public static ExportSettings Instance => _instance ??= Load();

        public int ImageWidth { get; set; } = 1280;
        public int ImageHeight { get; set; } = 720;
        public bool EnableCloudinary { get; set; } = true;
        public List<string> SubfolderOptions { get; set; } = new List<string> { "Open", "Close", "Internal" };

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? "");
                var lines = new List<string>
                {
                    $"ImageWidth={ImageWidth}",
                    $"ImageHeight={ImageHeight}",
                    $"EnableCloudinary={EnableCloudinary}",
                    $"SubfolderOptions={string.Join("|", SubfolderOptions)}"
                };
                File.WriteAllLines(SettingsPath, lines);
            }
            catch { }
        }

        private static ExportSettings Load()
        {
            var settings = new ExportSettings();
            try
            {
                if (File.Exists(SettingsPath))
                {
                    foreach (var line in File.ReadAllLines(SettingsPath))
                    {
                        var parts = line.Split('=');
                        if (parts.Length != 2) continue;

                        switch (parts[0])
                        {
                            case "ImageWidth":
                                if (int.TryParse(parts[1], out int w)) settings.ImageWidth = w;
                                break;
                            case "ImageHeight":
                                if (int.TryParse(parts[1], out int h)) settings.ImageHeight = h;
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
    [Plugin("DaabViewpointCreator", "DAAB",
    DisplayName = "Create Viewpoint",
    ToolTip = "Create organized viewpoint")]
    [AddInPlugin(AddInLocation.None)]
    public class ViewpointCreatorPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                var doc = NavisApplication.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No active document.", "Daab Viewpoint Creator");
                    return 0;
                }

                using (var dialog = new ViewpointCreatorForm(doc))
                {
                    var mainWindow = Process.GetCurrentProcess().MainWindowHandle;
                    if (mainWindow != IntPtr.Zero)
                    {
                        dialog.ShowDialog(new HwndWrapper(mainWindow));
                    }
                    else
                    {
                        dialog.ShowDialog();
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Daab Viewpoint Creator");
                return -1;
            }
        }
    }

    // Viewpoint creator dialog
    public class ViewpointCreatorForm : Form
    {
        private readonly Document _doc;
        private ComboBox categoryCombo = null!;
        private ComboBox levelCombo = null!;
        private ComboBox subfolderText = null!;
        private TextBox viewNameText = null!;
        private TextBox commentText = null!;
        private Button createButton = null!;
        private Button cancelButton = null!;
        private CheckBox addClashGeometryCheckbox = null!;

        public ViewpointCreatorForm(Document doc)
        {
            _doc = doc;
            InitializeComponents();
            LoadExistingStructure();
        }

        private void InitializeComponents()
        {
            Text = "Create Viewpoint";
            Width = 450;
            Height = 450;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            int y = 20;
            var labelCategory = new Label { Text = "Category:", Left = 20, Top = y, Width = 100 };
            categoryCombo = new ComboBox { Left = 130, Top = y, Width = 280, DropDownStyle = ComboBoxStyle.DropDown };

            y += 40;
            var labelLevel = new Label { Text = "Level:", Left = 20, Top = y, Width = 100 };
            levelCombo = new ComboBox { Left = 130, Top = y, Width = 280, DropDownStyle = ComboBoxStyle.DropDown };

            y += 40;
            var labelSubfolder = new Label { Text = "Subfolder:", Left = 20, Top = y, Width = 100 };
            subfolderText = new ComboBox { Left = 130, Top = y, Width = 280, DropDownStyle = ComboBoxStyle.DropDown };

            y += 40;
            var labelViewName = new Label { Text = "View Name:", Left = 20, Top = y, Width = 100 };
            viewNameText = new TextBox { Left = 130, Top = y, Width = 280 };

            y += 40;
            var labelComment = new Label { Text = "Comment:", Left = 20, Top = y, Width = 100 };
            commentText = new TextBox { Left = 130, Top = y, Width = 280, Height = 80, Multiline = true };

            y += 100;
            addClashGeometryCheckbox = new CheckBox { Text = "Add clash location geometry", Left = 20, Top = y, Width = 300 };

            y += 40;
            createButton = new Button { Text = "Create", Left = 230, Top = y, Width = 80 };
            createButton.Click += CreateButton_Click;

            cancelButton = new Button { Text = "Cancel", Left = 330, Top = y, Width = 80 };
            cancelButton.Click += CancelButton_Click;

            Controls.AddRange(new Control[] {
                labelCategory, categoryCombo, labelLevel, levelCombo,
                labelSubfolder, subfolderText, labelViewName, viewNameText,
                labelComment, commentText, addClashGeometryCheckbox,
                createButton, cancelButton
            });

            AcceptButton = createButton;
            CancelButton = cancelButton;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void LoadExistingStructure()
        {
            var root = _doc.SavedViewpoints?.RootItem;
            if (root == null) return;

            var categories = new HashSet<string>();
            var levels = new HashSet<string>();

            void TraverseForStructure(SavedItem item, int depth, string currentCategory)
            {
                if (item is GroupItem group)
                {
                    if (depth == 0) categories.Add(group.DisplayName ?? "");
                    else if (depth == 1) levels.Add(group.DisplayName ?? "");

                    foreach (SavedItem child in group.Children)
                        TraverseForStructure(child, depth + 1, depth == 0 ? (group.DisplayName ?? "") : currentCategory);
                }
            }

            foreach (SavedItem child in root.Children)
                TraverseForStructure(child, 0, "");

            categoryCombo.Items.AddRange(categories.OrderBy(x => x).ToArray());
            levelCombo.Items.AddRange(levels.OrderBy(x => x).ToArray());

            if (!categoryCombo.Items.Contains("CLASHES")) categoryCombo.Items.Add("CLASHES");
            if (!categoryCombo.Items.Contains("3D VIEWS")) categoryCombo.Items.Add("3D VIEWS");

            subfolderText.Items.Clear();
            foreach (var option in ExportSettings.Instance.SubfolderOptions)
                subfolderText.Items.Add(option);
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(viewNameText.Text))
            {
                MessageBox.Show("Please enter a view name.", "Validation");
                return;
            }

            try
            {
                var vp = new SavedViewpoint(_doc.CurrentViewpoint.CreateCopy());
                vp.DisplayName = viewNameText.Text;

                using (var transaction = _doc.Database.BeginTransaction(DatabaseChangedAction.Edited))
                {
                    var root = _doc.SavedViewpoints.RootItem;
                    var category = GetOrCreateFolder(root, categoryCombo.Text);
                    var level = string.IsNullOrWhiteSpace(levelCombo.Text) ? category : GetOrCreateFolder(category, levelCombo.Text);
                    var subfolder = string.IsNullOrWhiteSpace(subfolderText.Text) ? level : GetOrCreateFolder(level, subfolderText.Text);

                    subfolder.Children.Add(vp);

                    if (!string.IsNullOrWhiteSpace(commentText.Text))
                    {
                        AddComment(vp, commentText.Text);
                    }

                    transaction.Commit();
                }

                MessageBox.Show($"Viewpoint '{viewNameText.Text}' created successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating viewpoint: {ex.Message}\n\nDetails: {ex.StackTrace}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private GroupItem GetOrCreateFolder(GroupItem parent, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return parent;

            foreach (SavedItem item in parent.Children)
            {
                if (item is GroupItem group && group.DisplayName == name)
                    return group;
            }

            var newFolder = new FolderItem { DisplayName = name };
            parent.Children.Add(newFolder);
            return newFolder;
        }

        private void AddComment(SavedViewpoint vp, string commentText)
        {
            try
            {
                var commentsProperty = vp.GetType().GetProperty("Comments");
                if (commentsProperty != null)
                {
                    var comments = commentsProperty.GetValue(vp) as IList;
                    if (comments != null)
                    {
                        var commentType = comments.GetType().GetGenericArguments().FirstOrDefault();
                        if (commentType != null)
                        {
                            var comment = Activator.CreateInstance(commentType);
                            commentType.GetProperty("Body")?.SetValue(comment, commentText);
                            commentType.GetProperty("Status")?.SetValue(comment, "New");
                            commentType.GetProperty("Author")?.SetValue(comment, Environment.UserName);
                            commentType.GetProperty("CreationDate")?.SetValue(comment, DateTime.Now);
                            comments.Add(comment);
                        }
                    }
                }
            }
            catch
            {
                // Comment addition failed, but viewpoint was still created
            }
        }
    }

    // Main export plugin
    [Plugin("DaabNavisExport", "DAAB",
    DisplayName = "DaabReport",
    ToolTip = "Exports Navisworks viewpoints and comments to Daab Reports format")]
    [AddInPlugin(AddInLocation.None)]
    public class ExportPlugin : AddInPlugin
    {
        private const string DbFolderName = "DB";
        private const string ImagesFolderName = "Images";
        private const bool EnableSectioning = true;
        private const ImageGenerationStyle ExportStyle = ImageGenerationStyle.ScenePlusOverlay;

        // Cloudinary credentials
        private const string CloudinaryCloudName = "dhxutzg5f";
        private const string CloudinaryApiKey = "513751345586948";
        private const string CloudinaryApiSecret = "n3s3g_bCu7etUh6sRno9r8uQfNk";

        private static string _logFilePath = "";
        private static int _viewCounter = 0;
        private static int _totalViews = 0;
        private static int _processedViews = 0;
        private static Cloudinary? _cloudinary;
        private static string _projectName = "";

        public override int Execute(params string[] parameters)
        {
            try
            {
                if (NavisApplication.ActiveDocument == null)
                {
                    MessageBox.Show("No active document open.", "Daab Export");
                    return 0;
                }

                var document = NavisApplication.ActiveDocument;
                var settings = ExportSettings.Instance;

                var outputDirectory = ResolveOutputDirectory(parameters);
                var exportContext = BuildExportContext(document, outputDirectory);

                _projectName = Path.GetFileName(
                    exportContext.ProjectDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));

                _logFilePath = Path.Combine(exportContext.DbDirectory, "DaabExport.log");
                Log("=== Export started " + DateTime.Now.ToString("u") + " ===");
                Log($"Image resolution: {settings.ImageWidth}x{settings.ImageHeight}");

                // Initialize Cloudinary if enabled
                if (settings.EnableCloudinary)
                {
                    InitializeCloudinary();
                    Log("Cloudinary integration enabled");
                }
                else
                {
                    Log("Cloudinary integration disabled - using local paths only");
                }

                var root = document.SavedViewpoints?.RootItem;
                if (root == null)
                {
                    Log("No SavedViewpoints root found. Aborting.");
                    return 0;
                }

                _totalViews = CountSavedViews(root);
                _processedViews = 0;
                _viewCounter = 0;

                var csvPath = Path.Combine(exportContext.DbDirectory, "navisworks_views_comments.csv");
                using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    writer.WriteLine("Category,Level,Subfolder,ViewName,GUID,CommentID,Status,User,Body,CreatedDate,ImagePath,PublicImageURL");

                    bool cancelled = false;
                    Progress progress = null;
                    try
                    {
                        progress = NavisApplication.BeginProgress();

                        foreach (SavedItem child in root.Children)
                        {
                            TraverseAndExport(writer, document, exportContext, child, new List<string>(), progress, settings);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        cancelled = true;
                        Log("=== Export cancelled by user ===");
                    }
                    finally
                    {
                        NavisApplication.EndProgress();
                    }

                    if (!cancelled)
                    {
                        Log("CSV written: " + csvPath);
                        Log("=== Export finished successfully ===");

                        CopyPowerBITemplate(exportContext);

                        Process.Start(new ProcessStartInfo
                        {
                            FileName = exportContext.ProjectDirectory,
                            UseShellExecute = true
                        });
                    }

                    return 0;
                }
            }
            catch (Exception ex)
            {
                Log("Export failed: " + ex);
                MessageBox.Show("Export failed: " + ex.Message, "Daab Export");
                return -1;
            }
        }

        private static void InitializeCloudinary()
        {
            var account = new Account(CloudinaryCloudName, CloudinaryApiKey, CloudinaryApiSecret);
            _cloudinary = new Cloudinary(account);
        }

        private static void CopyPowerBITemplate(ExportContext exportContext)
        {
            try
            {
                var assemblyDir = Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";
                var src = Path.Combine(assemblyDir, "DaabReport.pbix");

                Log("PBIX lookup:");
                Log("  AssemblyDir: " + assemblyDir);
                Log("  Candidate  : " + src);

                if (!File.Exists(src))
                {
                    Log("DaabReport.pbix not found next to the plugin DLL.");
                    return;
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
                Log("Failed to copy Power BI template: " + ex.ToString());
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

            return new ExportContext(document, outputDirectory, projectDirectory, dbDirectory, imagesDirectory);
        }

        private static string ResolveProjectFolderName(Document document)
        {
            var sourceName = document.FileName;
            if (!string.IsNullOrWhiteSpace(sourceName))
            {
                var stem = Path.GetFileNameWithoutExtension(sourceName);
                return ToSafeFileName(stem);
            }
            return "Navisworks_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
        }

        private static string ResolveOutputDirectory(IReadOnlyList<string> parameters)
        {
            var explicitPath = parameters?.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
            if (!string.IsNullOrEmpty(explicitPath))
                return explicitPath;

            var myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(myDocs, "DaabNavisExport");
        }

        private static int CountSavedViews(SavedItem item)
        {
            int count = 0;
            if (item is GroupItem gi)
            {
                foreach (SavedItem child in gi.Children)
                    count += CountSavedViews(child);
            }
            else if (item is SavedViewpoint)
            {
                count++;
            }
            return count;
        }

        private static void TraverseAndExport(
            StreamWriter writer,
            Document doc,
            ExportContext ctx,
            SavedItem item,
            List<string> folderPath,
            Progress progress,
            ExportSettings settings)
        {
            if (item is GroupItem folder)
            {
                var nextPath = new List<string>(folderPath) { folder.DisplayName ?? "" };
                foreach (SavedItem child in folder.Children)
                    TraverseAndExport(writer, doc, ctx, child, nextPath, progress, settings);
                return;
            }

            if (item is not SavedViewpoint vp)
                return;

            string category = folderPath.Count >= 1 ? folderPath[0] : "";
            string level = folderPath.Count >= 2 ? folderPath[1] : "";
            string subfolder = "";
            if (folderPath.Count >= 3)
                subfolder = string.Join(" / ", folderPath.Skip(2).Where(s => !string.IsNullOrWhiteSpace(s)));

            _viewCounter++;
            string imageFile = $"vp{_viewCounter:D4}.jpg";
            string imageFullPath = Path.Combine(ctx.ImagesDirectory, imageFile);

            bool imageOk = TryGenerateImage(doc, vp, imageFullPath, settings.ImageWidth, settings.ImageHeight);
            if (!imageOk)
            {
                Log("Image render failed for: " + vp.DisplayName + " -> " + imageFullPath);
            }

            string publicUrl = "";
            if (imageOk && settings.EnableCloudinary && _cloudinary != null)
            {
                publicUrl = UploadToCloudinary(imageFullPath, vp.DisplayName ?? $"view_{_viewCounter}");
            }

            var comments = ExtractComments(vp);
            string viewName = vp.DisplayName ?? "";
            string guid = vp.Guid.ToString();

            if (comments.Count == 0)
            {
                WriteCsvRow(writer, category, level, subfolder, viewName, guid,
                            commentId: "", status: "", user: "", body: "", createdDate: "",
                            imageFile: imageFile, publicUrl: publicUrl);
            }
            else
            {
                foreach (var c in comments)
                {
                    WriteCsvRow(writer, category, level, subfolder, viewName, guid,
                                commentId: c.Id ?? "",
                                status: c.Status ?? "",
                                user: c.User ?? "",
                                body: c.Body ?? "",
                                createdDate: c.CreatedDate ?? "",
                                imageFile: imageFile, publicUrl: publicUrl);
                }
            }

            _processedViews = Math.Min(_processedViews + 1, _totalViews);
            double fraction = (_totalViews == 0) ? 1.0 : (_processedViews / (double)_totalViews);

            // Update progress bar and respect Cancel
            var cont = progress.Update(fraction);

            Log($"Exporting view {_processedViews}/{_totalViews}: {vp.DisplayName ?? imageFile}");

            if (!cont)
            {
                Log("Export cancelled by user at: " + (vp.DisplayName ?? "(unnamed)"));
                throw new OperationCanceledException("User cancelled export.");
            }
        }

        private static string UploadToCloudinary(string localPath, string viewName)
        {
            try
            {
                Log($"Uploading to Cloudinary: {viewName}");

                var safeViewName = ToSafeFileName(viewName);
                var uniqueId = $"{_projectName}_{safeViewName}";

                var uploadParams = new ImageUploadParams()
                {
                    File = new FileDescription(localPath),
                    PublicId = uniqueId,
                    Folder = "navisworks",
                    Overwrite = true,
                    UseFilename = false,
                    UniqueFilename = false
                };

                var uploadResult = _cloudinary!.Upload(uploadParams);

                if (uploadResult.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    Log($"Upload successful: {uploadResult.SecureUrl}");
                    return uploadResult.SecureUrl.ToString();
                }
                else
                {
                    Log($"Upload failed: {uploadResult.Error?.Message ?? "Unknown error"}");
                    return "";
                }
            }
            catch (Exception ex)
            {
                Log($"Cloudinary upload exception for {viewName}: {ex.Message}");
                return "";
            }
        }

        private static void WriteCsvRow(
            StreamWriter writer,
            string category,
            string level,
            string subfolder,
            string viewName,
            string guid,
            string commentId,
            string status,
            string user,
            string body,
            string createdDate,
            string imageFile,
            string publicUrl)
        {
            string San(string s) => (s ?? "").Replace(",", " ");

            writer.WriteLine(string.Join(",",
                San(category),
                San(level),
                San(subfolder),
                San(viewName),
                San(guid),
                San(commentId),
                San(status),
                San(user),
                San(body),
                San(createdDate),
                San(imageFile),
                San(publicUrl)
            ));
        }

        private static List<CommentInfo> ExtractComments(SavedViewpoint vp)
        {
            var list = new List<CommentInfo>();
            try
            {
                var commentsProp = vp.GetType().GetProperty("Comments");
                if (commentsProp?.GetValue(vp) is IEnumerable comments)
                {
                    foreach (var c in comments)
                    {
                        if (c == null) continue;
                        var t = c.GetType();

                        string id = t.GetProperty("Guid")?.GetValue(c)?.ToString();
                        string status = t.GetProperty("Status")?.GetValue(c)?.ToString();
                        string user = t.GetProperty("Author")?.GetValue(c)?.ToString();
                        string body = t.GetProperty("Body")?.GetValue(c)?.ToString();

                        string created = "";
                        var dtObj = t.GetProperty("CreationDate")?.GetValue(c);
                        if (dtObj is DateTime dt && dt.Year >= 1900)
                            created = dt.ToString("M/d/yyyy", CultureInfo.InvariantCulture);

                        list.Add(new CommentInfo
                        {
                            Id = id,
                            Status = status,
                            User = user,
                            Body = body,
                            CreatedDate = created
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Failed to reflect comments: " + ex.Message);
            }
            return list;
        }

        private static bool TryGenerateImage(Document doc, SavedViewpoint vp, string targetPath, int width, int height)
        {
            try
            {
                doc.SavedViewpoints.CurrentSavedViewpoint = vp;

                var view = doc.ActiveView;
                if (view == null)
                {
                    Log("ActiveView is null when exporting: " + (vp.DisplayName ?? "(unnamed)"));
                    return false;
                }

                using (var bmp = view.GenerateImage(ExportStyle, width, height, EnableSectioning))
                {
                    if (bmp == null)
                    {
                        Log("GenerateImage returned null for: " + (vp.DisplayName ?? "(unnamed)"));
                        return false;
                    }

                    var dir = Path.GetDirectoryName(targetPath);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                    bmp.Save(targetPath, ImageFormat.Jpeg);

                    bool exists = File.Exists(targetPath);
                    if (!exists)
                        Log("Image file not found after save: " + targetPath);

                    return exists;
                }
            }
            catch (Exception ex)
            {
                Log("Image export failed for " + (vp.DisplayName ?? "(unnamed)") + ": " + ex.Message);
                return false;
            }
        }

        private static void Log(string message)
        {
            try
            {
                var line = DateTime.Now.ToString("u") + " " + message + Environment.NewLine;
                Debug.WriteLine(line);
                if (!string.IsNullOrEmpty(_logFilePath))
                    File.AppendAllText(_logFilePath, line);
            }
            catch
            {
                // ignore logging errors
            }
        }

        private static string ToSafeFileName(string s)
        {
            if (string.IsNullOrEmpty(s)) return "Untitled";
            foreach (var ch in Path.GetInvalidFileNameChars())
                s = s.Replace(ch, '_');
            return s.Trim();
        }
    }

    internal sealed class ExportContext
    {
        public Document Document { get; }
        public string OutputDirectory { get; }
        public string ProjectDirectory { get; }
        public string DbDirectory { get; }
        public string ImagesDirectory { get; }

        public ExportContext(Document document, string outputDirectory, string projectDirectory, string dbDirectory, string imagesDirectory)
        {
            Document = document;
            OutputDirectory = outputDirectory;
            ProjectDirectory = projectDirectory;
            DbDirectory = dbDirectory;
            ImagesDirectory = imagesDirectory;
        }
    }

    internal sealed class CommentInfo
    {
        public string? Id { get; set; }
        public string? Status { get; set; }
        public string? User { get; set; }
        public string? Body { get; set; }
        public string? CreatedDate { get; set; }
    }
}
