using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Export
{
    public class CsvIntervalExporter : IIntervalExporter
    {
        public void Export(List<PriceInterval> intervals, List<string> sipps, double adjustmentFactor, double priceMultiplier, RateMonitorOptions options)
        {
            if (intervals.Count == 0)
            {
                Console.WriteLine("No intervals to export.");
                return;
            }

            Console.WriteLine("\nExporting intervals and rates to CSV...");

            var linkedSippGroups = new List<HashSet<string>>
            {
                new HashSet<string> { "EMMS", "PSMS" },
                new HashSet<string> { "PMMS", "PSAS" }
            };

            Console.WriteLine("\nLinked Sipp categories (will have the same prices):");
            foreach (var group in linkedSippGroups)
            {
                Console.WriteLine($"  {string.Join(" and ", group)}");
            }

            string outputDirectory = options.OutputDirectory;
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string adjustmentString = adjustmentFactor >= 0 ? $"plus{adjustmentFactor}" : $"minus{Math.Abs(adjustmentFactor)}";
            string csvFilePath = Path.Combine(outputDirectory, $"price_intervals_{timestamp}_{adjustmentString}.csv");

            try
            {
                using (var writer = new StreamWriter(csvFilePath))
                {
                    writer.Write("Sipp");
                    foreach (var interval in intervals)
                    {
                        writer.Write($",{interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}");
                    }
                    writer.WriteLine();

                    var adjustedPrices = new Dictionary<string, Dictionary<int, double>>();
                    foreach (var sipp in sipps)
                    {
                        adjustedPrices[sipp] = new Dictionary<int, double>();
                    }

                    for (int i = 0; i < intervals.Count; i++)
                    {
                        var interval = intervals[i];
                        foreach (var sipp in sipps)
                        {
                            if (interval.PricesBySipp.TryGetValue(sipp, out double price))
                            {
                                double adjustedPrice = price * priceMultiplier;
                                adjustedPrices[sipp][i] = adjustedPrice;
                            }
                        }
                    }

                    foreach (var linkedGroup in linkedSippGroups)
                    {
                        var selectedLinkedSipps = linkedGroup.Where(s => sipps.Contains(s)).ToList();
                        if (selectedLinkedSipps.Count <= 1)
                            continue;

                        for (int i = 0; i < intervals.Count; i++)
                        {
                            double highestPrice = 0;
                            foreach (var sipp in selectedLinkedSipps)
                            {
                                if (adjustedPrices[sipp].TryGetValue(i, out double price) && price > highestPrice)
                                {
                                    highestPrice = price;
                                }
                            }

                            if (highestPrice > 0)
                            {
                                foreach (var sipp in selectedLinkedSipps)
                                {
                                    adjustedPrices[sipp][i] = highestPrice;
                                }

                                Console.WriteLine($"Interval {i}: Set price for linked group {string.Join(", ", selectedLinkedSipps)} to {highestPrice:F2}");
                            }
                        }
                    }

                    foreach (var sipp in sipps.OrderBy(s => s))
                    {
                        writer.Write(sipp);

                        for (int i = 0; i < intervals.Count; i++)
                        {
                            if (adjustedPrices[sipp].TryGetValue(i, out double adjustedPrice))
                            {
                                writer.Write($",{adjustedPrice:F2}");
                            }
                            else
                            {
                                writer.Write(",N/A");
                            }
                        }
                        writer.WriteLine();
                    }

                    writer.WriteLine();
                    writer.WriteLine($"Prices adjusted by {adjustmentFactor}% ({priceMultiplier:F2}x multiplier)");
                    writer.WriteLine();
                    writer.WriteLine("Linked Sipp categories (same prices):");
                    foreach (var group in linkedSippGroups)
                    {
                        writer.WriteLine($"{string.Join(" and ", group)}");
                    }
                }

                Console.WriteLine($"Exported to: {csvFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to CSV: {ex.Message}");
            }
        }
    }
}
