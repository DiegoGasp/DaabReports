# Daab Navis Export

This folder contains a Navisworks 2026 add-in that automates the Daab Reports export workflow. The add-in performs the following actions when executed inside Autodesk Navisworks Manage or Simulate 2026:

1. Exports all saved viewpoints (and their folder hierarchy) from the active document to a Navisworks exchange XML file.
2. Runs the logic from `parseXml.py` (ported to C# in `Parsing/NavisworksXmlParser.cs`) to transform the XML into the `navisworks_views_comments.csv` structure used by the Power BI template.
3. Writes a `debug.txt` log that mirrors the diagnostics produced by the Python tooling.
4. Exports viewport images (JPEG) whose filenames align with the `ImagePath` column produced by the parser.

The resulting files are written beneath `%USERPROFILE%\Documents\DaabNavisExport\<Project Name>` by default. You can supply a different output directory by passing a path parameter in the Navisworks **Add-Ins** window when you launch the plugin.

## Project layout

```
DaabNavisExport/
├── DaabNavisExport.csproj          # .NET Framework 4.8 class library project
├── ExportPlugin.cs                 # Add-in entry point (`ExportPlugin`)
├── Parsing/
│   ├── NavisworksXmlParser.cs      # C# port of parseXml.py
│   └── ParseResult.cs              # Parser result container
├── Properties/
│   └── AssemblyInfo.cs             # Assembly metadata
├── Utilities/
│   ├── ExportContext.cs            # Export session state
│   └── PathSanitizer.cs            # File/Path helper utilities
└── parseXml.py                     # Original Python script for reference
```

## Building

1. Open the solution folder in Visual Studio 2022.
2. Add references to the Navisworks 2026 API assemblies:
   - `Autodesk.Navisworks.Api.dll`
   - `Autodesk.Navisworks.Api.DocumentParts.dll`
   These are typically located in `C:\Program Files\Autodesk\Navisworks Manage 2026\api\`. Ensure the references are set to **Copy Local = false**.
3. Build the project in **Release** mode. The output `DaabNavisExport.dll` will be placed in `bin/Release`.

## Deployment

1. Create an add-in bundle folder (e.g. `%APPDATA%\Autodesk\Navisworks Manage 2026\Plugins\DaabNavisExport.bundle`).
2. Inside the bundle, add the compiled `DaabNavisExport.dll`, the `parseXml.py` reference file (optional), and a `PackageContents.xml` similar to the following:

```xml
<?xml version="1.0" encoding="utf-8"?>
<ApplicationPackage SchemaVersion="1.0" AutodeskProduct="Navisworks" ProductType="Application" Name="Daab Navis Export" Description="Exports viewpoints, comments, and images to Daab Reports format." AppVersion="1.0" ProductCode="{E324E173-2803-489B-B727-34A96E616D67}" UpgradeCode="{31D68667-ED13-4805-B5D7-3E06D814AF03}">
  <CompanyDetails Name="Daab Reports" Url="https://daabreports.example"/>
  <Components>
    <RuntimeRequirements SeriesMin="2026" SeriesMax="2026"/>
    <ComponentEntry AppName="DaabNavisExport" Version="1.0" ModuleName="DaabNavisExport.dll" AppType="Application" LoadOnStartUp="True"/>
  </Components>
</ApplicationPackage>
```

3. Launch Navisworks 2026 and open the **Add-Ins** tab. You should find **Daab Navis Export** listed. Running it will produce the following structure (matching the sample project layout):

   ```
   <Project Name>/
   ├── DB/
   │   ├── DB.xml
   │   ├── debug.txt
   │   └── navisworks_views_comments.csv
   └── Images/
       ├── vp0001.jpg
       ├── vp0002.jpg
       └── ...
   ```

## Notes

- The parser honours the same duplicate filtering, logging format, and date parsing rules as the original Python implementation in `1717 N FLAG/DB/ParseXml.py`.
- Image thumbnails are generated at 800×450 resolution; adjust `GenerateThumbnail` parameters in `ExportPlugin.cs` if a different size is required.
- If the add-in is executed without any saved viewpoints, the XML and CSV will still be generated but remain empty aside from headers.
