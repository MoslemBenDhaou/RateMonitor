using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using RateMonitorAnalyser.Logic.Analysis;
using RateMonitorAnalyser.Logic.Export;
using RateMonitorAnalyser.Logic.Interfaces;
using RateMonitorAnalyser.Logic.Models;
using RateMonitorAnalyser.Logic.Presentation;
using RateMonitorAnalyser.Logic.Selection;

namespace RateMonitorAnalyser.Logic
{
    public class RateMonitorRunner
    {
        private readonly RateMonitorOptions _options;
        private readonly IRateDataExtractor _extractor;
        private readonly ISippSelector _selector;
        private readonly IPriceIntervalAnalyzer _analyzer;
        private readonly IntervalPresenter _presenter;
        private readonly IEnumerable<IIntervalExporter> _exporters;

        public RateMonitorRunner(
            RateMonitorOptions options,
            IRateDataExtractor extractor,
            ISippSelector selector,
            IPriceIntervalAnalyzer analyzer,
            IntervalPresenter presenter,
            IEnumerable<IIntervalExporter> exporters)
        {
            _options = options;
            _extractor = extractor;
            _selector = selector;
            _analyzer = analyzer;
            _presenter = presenter;
            _exporters = exporters;
        }

        public void RunInteractive()
        {
            ExcelPackage.License.SetNonCommercialPersonal(_options.LicenseName);

            Console.WriteLine("RateMonitor Analyser Starting...");

            string sourceDirectory = _options.SourceDirectory;

            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                Console.WriteLine("Creating source directory...");
                Directory.CreateDirectory(sourceDirectory);
                Console.WriteLine("Source directory created. Please place files in this directory and run the program again.");
                return;
            }

            Regex regex = new Regex(_options.FilePattern);

            var matchingFiles = Directory.GetFiles(sourceDirectory)
                .Where(file => regex.IsMatch(Path.GetFileName(file)))
                .ToList();

            if (matchingFiles.Count == 0)
            {
                Console.WriteLine("No matching files found in the source directory.");
                Console.WriteLine("Expected file name format: suggestion_report_90468_2025-04-07.xlsx");
                return;
            }

            var mostRecentFile = matchingFiles
                .OrderByDescending(file => File.GetLastWriteTime(file))
                .First();

            Console.WriteLine($"Found most recent file: {Path.GetFileName(mostRecentFile)}");
            Console.WriteLine($"File path: {mostRecentFile}");
            Console.WriteLine($"Last modified: {File.GetLastWriteTime(mostRecentFile)}");

            var rateData = _extractor.ExtractRateData(mostRecentFile);

            var allSipps = rateData
                .Where(r => r.SuggestedAmount > 0)
                .Select(r => r.Sipp)
                .Distinct()
                .OrderBy(s => s)
                .ToList();

            var selectedSipps = _selector.PromptForSippSelection(allSipps);

            var intervals = _analyzer.AnalyzePriceIntervals(rateData, selectedSipps);

            Console.WriteLine("\nEnter price adjustment factor (e.g., 15 for 15% increase, -10 for 10% decrease, or 0 for no change):");
            string input = Console.ReadLine() ?? "0";

            double adjustmentFactor = 0;
            if (!double.TryParse(input, out adjustmentFactor))
            {
                Console.WriteLine("Invalid input. Using 0% adjustment (no change).");
                adjustmentFactor = 0;
            }

            double priceMultiplier = 1 + (adjustmentFactor / 100);

            Console.WriteLine($"Applying {adjustmentFactor}% adjustment to all prices (multiplier: {priceMultiplier:F2})");

            foreach (var exporter in _exporters)
            {
                exporter.Export(intervals, selectedSipps, adjustmentFactor, priceMultiplier, _options);
            }

            _presenter.DisplayIntervalsAndPromptForDetails(intervals);

            Console.WriteLine("Processing complete.");
        }
    }
}
