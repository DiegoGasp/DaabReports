using System.Collections.Generic;
using Autodesk.Navisworks.Api;

namespace DaabNavisExport.Utilities
{
    internal sealed class ExportContext
    {
        public ExportContext(Document document, string outputDirectory, string xmlPath)
        {
            Document = document;
            OutputDirectory = outputDirectory;
            XmlPath = xmlPath;
            ViewSequence = new List<SavedViewpoint>();
        }

        public Document Document { get; }

        public string OutputDirectory { get; }

        public string XmlPath { get; }

        public List<SavedViewpoint> ViewSequence { get; }
    }
}
