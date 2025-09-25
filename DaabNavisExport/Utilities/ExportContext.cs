using System.Collections.Generic;
using Autodesk.Navisworks.Api;

namespace DaabNavisExport.Utilities
{
    internal sealed class ExportContext
    {
        public ExportContext(
            Document document,
            string rootDirectory,
            string projectDirectory,
            string dbDirectory,
            string imagesDirectory,
            string xmlPath)
        {
            Document = document;
            RootDirectory = rootDirectory;
            ProjectDirectory = projectDirectory;
            DbDirectory = dbDirectory;
            ImagesDirectory = imagesDirectory;
            XmlPath = xmlPath;
            ViewSequence = new List<SavedViewpoint>();
        }

        public Document Document { get; }

        public string RootDirectory { get; }

        public string ProjectDirectory { get; }

        public string DbDirectory { get; }

        public string ImagesDirectory { get; }

        public string XmlPath { get; }

        public List<SavedViewpoint> ViewSequence { get; }
    }

}


