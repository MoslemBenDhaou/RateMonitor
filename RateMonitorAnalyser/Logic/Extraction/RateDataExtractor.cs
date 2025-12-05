using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using RateMonitorAnalyser.Logic.Interfaces;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Extraction
{
    public class RateDataExtractor : IRateDataExtractor
    {
        public List<RateInfo> ExtractRateData(string filePath)
        {
            Console.WriteLine("\nExtracting rate data...");
            var rateData = new List<RateInfo>();

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        Console.WriteLine("No worksheet found in the Excel file.");
                        return rateData;
                    }

                    int rows = worksheet.Dimension?.Rows ?? 0;
                    int columns = worksheet.Dimension?.Columns ?? 0;

                    int sippColIndex = -1;
                    int pickupDateColIndex = -1;
                    int suggestedAmountColIndex = -1;
                    int ruleDescriptionColIndex = -1;
                    int locationColIndex = -1;
                    int lorColIndex = -1;

                    for (int col = 1; col <= columns; col++)
                    {
                        var headerValue = worksheet.Cells[1, col].Value?.ToString();
                        if (headerValue == "Sipp") sippColIndex = col;
                        else if (headerValue == "PickUpDate") pickupDateColIndex = col;
                        else if (headerValue == "SuggestedAmount") suggestedAmountColIndex = col;
                        else if (headerValue == "RuleDescription") ruleDescriptionColIndex = col;
                        else if (headerValue == "Location") locationColIndex = col;
                        else if (headerValue == "Lor") lorColIndex = col;
                    }

                    if (sippColIndex == -1 || pickupDateColIndex == -1 || suggestedAmountColIndex == -1)
                    {
                        Console.WriteLine("Required columns not found in the Excel file.");
                        return rateData;
                    }

                    for (int row = 2; row <= rows; row++)
                    {
                        var sipp = worksheet.Cells[row, sippColIndex].Value?.ToString();
                        var pickupDateValue = worksheet.Cells[row, pickupDateColIndex].Value;
                        var suggestedAmountValue = worksheet.Cells[row, suggestedAmountColIndex].Value;
                        var ruleDescription = worksheet.Cells[row, ruleDescriptionColIndex].Value?.ToString() ?? "";
                        var location = worksheet.Cells[row, locationColIndex].Value?.ToString() ?? "";
                        var lorValue = worksheet.Cells[row, lorColIndex].Value;

                        if (string.IsNullOrEmpty(sipp) || pickupDateValue == null)
                            continue;

                        double suggestedAmount = 0;
                        if (suggestedAmountValue != null)
                        {
                            double.TryParse(suggestedAmountValue.ToString(), out suggestedAmount);
                            suggestedAmount = Math.Round(suggestedAmount * 10) / 10;
                        }

                        DateTime pickupDate;
                        if (pickupDateValue is double excelDate)
                        {
                            pickupDate = DateTime.FromOADate(excelDate);
                        }
                        else if (!DateTime.TryParse(pickupDateValue.ToString(), out pickupDate))
                        {
                            continue;
                        }

                        int lor = 0;
                        if (lorValue != null)
                        {
                            int.TryParse(lorValue.ToString(), out lor);
                        }

                        rateData.Add(new RateInfo
                        {
                            PickupDate = pickupDate,
                            Sipp = sipp,
                            SuggestedAmount = suggestedAmount,
                            RuleDescription = ruleDescription,
                            Location = location,
                            Lor = lor
                        });
                    }
                }

                Console.WriteLine($"Extracted {rateData.Count} rate entries (including entries with 0 suggested amount).");

                rateData = rateData.OrderBy(r => r.PickupDate).ToList();

                if (rateData.Any())
                {
                    var firstDate = rateData.First().PickupDate;
                    var lastDate = rateData.Last().PickupDate;
                    Console.WriteLine($"Date range: {firstDate:yyyy-MM-dd} to {lastDate:yyyy-MM-dd} ({(lastDate - firstDate).TotalDays + 1} days)");

                    var uniqueDates = rateData.Select(r => r.PickupDate.Date).Distinct().Count();
                    Console.WriteLine($"Number of unique dates: {uniqueDates}");

                    var uniqueSipps = rateData.Select(r => r.Sipp).Distinct().Count();
                    Console.WriteLine($"Number of unique Sipp codes: {uniqueSipps}");

                    var zeroAmountCount = rateData.Count(r => r.SuggestedAmount == 0);
                    Console.WriteLine($"Number of entries with 0 suggested amount: {zeroAmountCount}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting rate data: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            return rateData;
        }
    }
}
