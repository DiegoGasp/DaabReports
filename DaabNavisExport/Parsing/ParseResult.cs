using System.Collections.Generic;

namespace DaabNavisExport.Parsing
{
    internal sealed class ParseResult
    {
        public ParseResult(IReadOnlyList<List<string?>> rows, IReadOnlyList<string> debugLines)
        {
            var projected = new List<IReadOnlyList<string?>>(rows.Count);
            foreach (var row in rows)
            {
                projected.Add(row);
            }

            Rows = projected;
            DebugLines = debugLines;
        }

        public IReadOnlyList<IReadOnlyList<string?>> Rows { get; }

        public IReadOnlyList<string> DebugLines { get; }
    }
}
