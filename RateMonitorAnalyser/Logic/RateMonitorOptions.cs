using System.IO;

namespace RateMonitorAnalyser.Logic
{
    public class RateMonitorOptions
    {
        public string WorkingDirectory { get; set; } = Directory.GetCurrentDirectory();

        public string SourceDirectory => Path.Combine(WorkingDirectory, "source");

        public string OutputDirectory => Path.Combine(WorkingDirectory, "output");

        public string ExportDirectory => Path.Combine(WorkingDirectory, "Export");

        public string FilePattern { get; set; } = @"suggestion_report_\d+_\d{4}-\d{2}-\d{2}\.xlsx";

        public string ReferenceSipp { get; set; } = "ESMS";

        public string LicenseName { get; set; } = "RateMonitorAnalyser";
    }
}
