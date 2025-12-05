using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Export
{
    public class ExcelTemplateExporter : IIntervalExporter
    {
        public void Export(List<PriceInterval> intervals, List<string> sipps, double adjustmentFactor, double priceMultiplier, RateMonitorOptions options)
        {
            Console.WriteLine("\nPreparing Excel template for export...");

            string exportDirectory = options.ExportDirectory;
            string templatePath = Path.Combine(exportDirectory, "TUN pricing template.xlsx");

            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            string todayDate = DateTime.Now.ToString("yyyy-MM-dd");
            string newFileName = $"TUN Pricing {todayDate}.xlsx";
            string newFilePath = Path.Combine(exportDirectory, newFileName);

            string referenceSipp = options.ReferenceSipp;

            if (!sipps.Contains(referenceSipp))
            {
                Console.WriteLine($"Warning: Reference Sipp {referenceSipp} not found in selected Sipps. Using first available Sipp instead.");
                referenceSipp = sipps.FirstOrDefault() ?? "N/A";
            }

            Console.WriteLine($"Using {referenceSipp} as reference Sipp for row 8 pricing.");

            var linkedSippGroups = new List<HashSet<string>>
            {
                // Intentionally empty to mirror original template export behavior
            };

            Console.WriteLine("\nLinked Sipp categories (will have the same prices):");
            foreach (var group in linkedSippGroups)
            {
                Console.WriteLine($"  {string.Join(" and ", group)}");
            }

            try
            {
                File.Copy(templatePath, newFilePath, true);
                Console.WriteLine($"Created new Excel file: {newFileName}");

                string sourceDirectory = options.SourceDirectory;
                var xdFiles = Directory.GetFiles(sourceDirectory, "*D_*.xlsx")
                    .Where(f => Regex.IsMatch(Path.GetFileName(f), @"^\d+D_.*\.xlsx$"))
                    .ToList();

                Console.WriteLine($"Found {xdFiles.Count} XD files for reserve length factors:");
                foreach (var file in xdFiles)
                {
                    Console.WriteLine($"  {Path.GetFileName(file)}");
                }

                var xdData = ExtractXDData(xdFiles, referenceSipp);

                using (var package = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets["PRICING"];

                    if (worksheet == null)
                    {
                        Console.WriteLine("PRICING worksheet not found in template.");
                        return;
                    }

                    Console.WriteLine("Found PRICING worksheet in template.");

                    var normalizedPrices = NormalizePricesForLinkedSipps(intervals, linkedSippGroups);

                    var sortedIntervals = intervals.OrderBy(i => i.StartDate).ToList();

                    var referenceSippRates = new Dictionary<int, Dictionary<DateTime, double>>();

                    for (int i = 0; i < sortedIntervals.Count; i++)
                    {
                        int columnIndex = 2 + (i * 2);
                        string columnLetter = GetExcelColumnName(columnIndex);

                        string startDateCell = $"{columnLetter}6";
                        worksheet.Cells[startDateCell].Value = sortedIntervals[i].StartDate;
                        worksheet.Cells[startDateCell].Style.Numberformat.Format = "yyyy-mm-dd";

                        string endDateCell = $"{columnLetter}7";
                        worksheet.Cells[endDateCell].Value = sortedIntervals[i].EndDate;
                        worksheet.Cells[endDateCell].Style.Numberformat.Format = "yyyy-mm-dd";

                        string rateCell = $"{columnLetter}8";

                        if (normalizedPrices.TryGetValue(i, out var pricesBySipp) &&
                            pricesBySipp.TryGetValue(referenceSipp, out double price))
                        {
                            double sevenDayRate = price * 7;
                            double adjustedSevenDayRate = sevenDayRate * priceMultiplier;

                            worksheet.Cells[rateCell].Value = adjustedSevenDayRate;
                            worksheet.Cells[rateCell].Style.Numberformat.Format = "#,##0.00";

                            if (!referenceSippRates.ContainsKey(i))
                            {
                                referenceSippRates[i] = new Dictionary<DateTime, double>();
                            }

                            for (DateTime date = sortedIntervals[i].StartDate; date <= sortedIntervals[i].EndDate; date = date.AddDays(1))
                            {
                                referenceSippRates[i][date] = price;
                            }

                            Console.WriteLine($"Filled interval {i + 1}: {columnLetter}6 = {sortedIntervals[i].StartDate:yyyy-MM-dd}, " +
                                             $"{columnLetter}7 = {sortedIntervals[i].EndDate:yyyy-MM-dd}, " +
                                             $"{columnLetter}8 = {adjustedSevenDayRate:F2} ({referenceSipp} 7-day rate)");
                        }
                        else
                        {
                            worksheet.Cells[rateCell].Value = "N/A";
                            Console.WriteLine($"Filled interval {i + 1}: {columnLetter}6 = {sortedIntervals[i].StartDate:yyyy-MM-dd}, " +
                                             $"{columnLetter}7 = {sortedIntervals[i].EndDate:yyyy-MM-dd}, " +
                                             $"{columnLetter}8 = N/A (No price data for {referenceSipp})");
                        }
                    }

                    Console.WriteLine("All intervals filled in PRICING sheet.");

                    int nextRow = 9;
                    bool hasMoreRows = true;

                    while (hasMoreRows)
                    {
                        var sippCell = worksheet.Cells[$"A{nextRow}"];
                        string? sipp = sippCell.Value?.ToString();

                        if (string.IsNullOrWhiteSpace(sipp))
                        {
                            hasMoreRows = false;
                            continue;
                        }

                        Console.WriteLine($"Found Sipp code in A{nextRow}: {sipp}");

                        for (int i = 0; i < sortedIntervals.Count; i++)
                        {
                            int rateColumnIndex = 2 + (i * 2);
                            string? rateColumnLetter = GetExcelColumnName(rateColumnIndex);

                            int percentageColumnIndex = 3 + (i * 2);
                            string? percentageColumnLetter = GetExcelColumnName(percentageColumnIndex);

                            if (rateColumnLetter != null && percentageColumnLetter != null)
                            {
                                string rateCell = $"{rateColumnLetter}{nextRow}";
                                string percentageCell = $"{percentageColumnLetter}{nextRow}";

                                if (referenceSippRates.TryGetValue(i, out var ratesForDates) &&
                                    normalizedPrices.TryGetValue(i, out var pricesBySipp))
                                {
                                    double referenceRate = ratesForDates[sortedIntervals[i].StartDate] * 7;

                                    double sippPrice;
                                    double sippMultiplier = 1.0;

                                    if (sipp == "PMAS")
                                    {
                                        sippMultiplier = 3.0;
                                        sippPrice = referenceRate * sippMultiplier;
                                        Console.WriteLine($"Applied special multiplier for {sipp}: {sippMultiplier}x");
                                    }
                                    else if (sipp == "SMAS")
                                    {
                                        sippMultiplier = 4.0;
                                        sippPrice = referenceRate * sippMultiplier;
                                        Console.WriteLine($"Applied special multiplier for {sipp}: {sippMultiplier}x");
                                    }
                                    else if (pricesBySipp.TryGetValue(sipp, out double price))
                                    {
                                        sippPrice = price * 7;
                                    }
                                    else
                                    {
                                        worksheet.Cells[rateCell].Value = "0";
                                        worksheet.Cells[percentageCell].Value = 0;
                                        worksheet.Cells[percentageCell].Style.Numberformat.Format = "0.0%";
                                        Console.WriteLine($"No price available for {sipp} in interval {i + 1}");
                                        continue;
                                    }

                                    double sippSevenDayRate = sippPrice;

                                    double percentageDiff = 0;
                                    if (referenceRate > 0)
                                    {
                                        percentageDiff = ((sippSevenDayRate - referenceRate) / referenceRate) * 100;
                                    }

                                    double adjustedReferenceRate = referenceRate * priceMultiplier;
                                    double adjustedSippSevenDayRate = sippSevenDayRate * priceMultiplier;

                                    worksheet.Cells[rateCell].Value = adjustedSippSevenDayRate;
                                    worksheet.Cells[rateCell].Style.Numberformat.Format = "#,##0.00";

                                    worksheet.Cells[percentageCell].Value = percentageDiff / 100;
                                    worksheet.Cells[percentageCell].Style.Numberformat.Format = "0.0%";

                                    Console.WriteLine($"Set {sipp} price for interval {i + 1}: {adjustedSippSevenDayRate:F2} in cell {rateCell}");
                                    Console.WriteLine($"Set {sipp} percentage difference to {referenceSipp} for interval {i + 1}: {percentageDiff:F1}% in cell {percentageCell}");
                                    Console.WriteLine($"  Original prices: {sipp}={sippSevenDayRate:F2}, {referenceSipp}={referenceRate:F2}");
                                    Console.WriteLine($"  Adjusted prices: {sipp}={adjustedSippSevenDayRate:F2}, {referenceSipp}={adjustedReferenceRate:F2}");
                                }
                                else
                                {
                                    worksheet.Cells[rateCell].Value = "N/A";
                                    worksheet.Cells[percentageCell].Value = 0;
                                    worksheet.Cells[percentageCell].Style.Numberformat.Format = "0.0%";
                                    Console.WriteLine($"Could not calculate price and percentage for {sipp} in interval {i + 1} - missing data");
                                }
                            }
                        }

                        nextRow++;

                        if (nextRow > 100)
                        {
                            Console.WriteLine("Reached maximum row limit (100). Stopping.");
                            break;
                        }
                    }

                    package.Save();
                    Console.WriteLine($"Saved Excel template with PRICING sheet data to {templatePath}");

                    var reserveSheet = package.Workbook.Worksheets["RESERVE LENGTH"];
                    if (reserveSheet == null)
                    {
                        Console.WriteLine("RESERVE LENGTH worksheet not found in template.");
                    }
                    else
                    {
                        Console.WriteLine("Found RESERVE LENGTH worksheet. Filling reserve length factors...");

                        var rentalLengthToColumn = new Dictionary<int, string?>
                        {
                            { 1, "B" },
                            { 2, "C" },
                            { 3, "D" },
                            { 4, "E" },
                            { 5, "F" },
                            { 6, "G" }
                        };

                        var nextRowForRentalLength = new Dictionary<int, int>();

                        foreach (var rentalLength in rentalLengthToColumn.Keys)
                        {
                            string? column = rentalLengthToColumn[rentalLength];
                            if (column == null) continue;

                            nextRowForRentalLength[rentalLength] = 2;
                            Console.WriteLine($"Starting row for {rentalLength}D (column {column}): 2");
                        }

                        const int rowSpacingBetweenIntervals = 13;

                        foreach (var interval in sortedIntervals)
                        {
                            Console.WriteLine($"Processing interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}");

                            for (int rentalLength = 1; rentalLength <= 6; rentalLength++)
                            {
                                if (rentalLengthToColumn.TryGetValue(rentalLength, out string? column) && column != null)
                                {
                                    double factor = CalculateReserveLengthFactor(
                                        interval,
                                        rentalLength,
                                        referenceSipp,
                                        xdData,
                                        referenceSippRates);

                                    if (factor > 0 && factor < 1)
                                    {
                                        Console.WriteLine($"  Adjusting factor from {factor:F3} to 1.000 (minimum allowed value)");
                                        factor = 1.0;
                                    }

                                    factor = Math.Ceiling(factor * 100) / 100;
                                    Console.WriteLine($"  Rounded factor to {factor:F3}");

                                    if (factor > 0)
                                    {
                                        if (nextRowForRentalLength.TryGetValue(rentalLength, out int rowNum))
                                        {
                                            string factorCell = $"{column}{rowNum}";
                                            reserveSheet.Cells[factorCell].Value = factor;
                                            reserveSheet.Cells[factorCell].Style.Numberformat.Format = "0.00";

                                            Console.WriteLine($"Set {rentalLength}D factor for interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}: {factor:F3} in cell {factorCell}");

                                            nextRowForRentalLength[rentalLength] = rowNum + rowSpacingBetweenIntervals;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"No factor calculated for {rentalLength}D, interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}");
                                    }
                                }
                            }
                        }
                    }

                    package.Save();
                    Console.WriteLine($"Saved Excel template with PRICING sheet data to {templatePath}");

                    package.Save();
                }

                Console.WriteLine($"Excel template prepared and saved to: {newFilePath}");
                Console.WriteLine($"Filled intervals in the PRICING sheet with dates, {referenceSipp} 7-day rates, and percentage differences.");
                Console.WriteLine("Linked Sipp groups were respected in the calculations.");
                Console.WriteLine("RESERVE LENGTH sheet was filled with factors for shorter rental lengths.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error preparing Excel template: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static Dictionary<int, Dictionary<DateTime, double>> ExtractXDData(List<string> xdFiles, string referenceSipp)
        {
            var result = new Dictionary<int, Dictionary<DateTime, double>>();

            foreach (var file in xdFiles)
            {
                string fileName = Path.GetFileName(file);
                var match = Regex.Match(fileName, @"^(\d+)D_");

                if (match.Success && int.TryParse(match.Groups[1].Value, out int rentalLength))
                {
                    Console.WriteLine($"Processing {fileName} for {rentalLength}D rental length");

                    try
                    {
                        using (var package = new ExcelPackage(new FileInfo(file)))
                        {
                            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                            if (worksheet != null)
                            {
                                result[rentalLength] = new Dictionary<DateTime, double>();

                                int sippColumn = -1;
                                int pickUpDateColumn = -1;
                                int suggestedAmountColumn = -1;

                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    string header = worksheet.Cells[1, col].Value?.ToString() ?? string.Empty;
                                    if (header == "Sipp")
                                        sippColumn = col;
                                    else if (header == "PickUpDate")
                                        pickUpDateColumn = col;
                                    else if (header == "SuggestedAmount")
                                        suggestedAmountColumn = col;
                                }

                                if (sippColumn > 0 && pickUpDateColumn > 0 && suggestedAmountColumn > 0)
                                {
                                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        string sipp = worksheet.Cells[row, sippColumn].Value?.ToString() ?? string.Empty;

                                        if (sipp == referenceSipp)
                                        {
                                            DateTime pickUpDate;

                                            if (worksheet.Cells[row, pickUpDateColumn].Value is DateTime dateValue)
                                            {
                                                pickUpDate = dateValue;
                                            }
                                            else if (worksheet.Cells[row, pickUpDateColumn].Value is double numericDate)
                                            {
                                                pickUpDate = DateTime.FromOADate(numericDate);
                                            }
                                            else if (worksheet.Cells[row, pickUpDateColumn].Value is int intDate)
                                            {
                                                pickUpDate = DateTime.FromOADate(intDate);
                                            }
                                            else if (worksheet.Cells[row, pickUpDateColumn].Value is string dateStr &&
                                                     DateTime.TryParse(dateStr, out DateTime parsedDate))
                                            {
                                                pickUpDate = parsedDate;
                                            }
                                            else
                                            {
                                                continue;
                                            }

                                            double suggestedAmount;
                                            if (worksheet.Cells[row, suggestedAmountColumn].Value is double doubleAmount)
                                            {
                                                suggestedAmount = doubleAmount;
                                            }
                                            else if (worksheet.Cells[row, suggestedAmountColumn].Value is int intAmount)
                                            {
                                                suggestedAmount = intAmount;
                                            }
                                            else if (worksheet.Cells[row, suggestedAmountColumn].Value is string amountStr &&
                                                     double.TryParse(amountStr, out double parsedAmount))
                                            {
                                                suggestedAmount = parsedAmount;
                                            }
                                            else
                                            {
                                                continue;
                                            }

                                            suggestedAmount = Math.Round(suggestedAmount * 10) / 10;

                                            result[rentalLength][pickUpDate] = suggestedAmount;

                                            Console.WriteLine($"  {rentalLength}D price for {pickUpDate:yyyy-MM-dd}: {suggestedAmount:F1}");
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"  Could not find required columns in {fileName}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Error processing {fileName}: {ex.Message}");
                    }
                }
            }

            return result;
        }

        private static double CalculateReserveLengthFactor(
            PriceInterval interval,
            int rentalLength,
            string referenceSipp,
            Dictionary<int, Dictionary<DateTime, double>> xdData,
            Dictionary<int, Dictionary<DateTime, double>> referenceSippRates)
        {
            Console.WriteLine($"Calculating factor for interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}, rental length {rentalLength}D");

            int actualRentalLength = rentalLength;
            while (actualRentalLength <= 6)
            {
                if (xdData.ContainsKey(actualRentalLength))
                    break;

                Console.WriteLine($"  No data for {actualRentalLength}D, trying {actualRentalLength + 1}D");
                actualRentalLength++;
            }

            if (actualRentalLength > 6 || !xdData.ContainsKey(actualRentalLength))
            {
                Console.WriteLine($"  No data available for any rental length >= {rentalLength}D");
                return 0;
            }

            if (actualRentalLength != rentalLength)
            {
                Console.WriteLine($"  Using {actualRentalLength}D data for {rentalLength}D factor");
            }

            if (!xdData[actualRentalLength].TryGetValue(interval.StartDate, out double xdPrice))
            {
                var closestDate = xdData[actualRentalLength].Keys
                    .OrderBy(d => Math.Abs((d - interval.StartDate).TotalDays))
                    .FirstOrDefault();

                if (closestDate != default)
                {
                    xdPrice = xdData[actualRentalLength][closestDate];
                    Console.WriteLine($"  Using closest date {closestDate:yyyy-MM-dd} for {actualRentalLength}D price: {xdPrice:F1}");
                }
                else
                {
                    Console.WriteLine($"  No price data found for {actualRentalLength}D rental length");
                    return 0;
                }
            }
            else
            {
                Console.WriteLine($"  Found exact price for {actualRentalLength}D on {interval.StartDate:yyyy-MM-dd}: {xdPrice:F1}");
            }

            int intervalIndex = -1;

            foreach (var entry in referenceSippRates)
            {
                if (entry.Value.ContainsKey(interval.StartDate))
                {
                    intervalIndex = entry.Key;
                    Console.WriteLine($"  Found interval index by start date: {intervalIndex}");
                    break;
                }
            }

            if (intervalIndex == -1)
            {
                Console.WriteLine($"  Could not find exact interval by start date, trying to find containing interval");

                for (int i = 0; i < referenceSippRates.Count; i++)
                {
                    if (referenceSippRates.TryGetValue(i, out var intervalRates))
                    {
                        var dates = intervalRates.Keys.OrderBy(d => d).ToList();

                        if (dates.Count > 0)
                        {
                            if (interval.StartDate >= dates.First() && interval.StartDate <= dates.Last())
                            {
                                intervalIndex = i;
                                Console.WriteLine($"  Found containing interval index: {intervalIndex}");
                                break;
                            }
                        }
                    }
                }
            }

            if (intervalIndex == -1)
            {
                Console.WriteLine($"  Could not find interval index for {interval.StartDate:yyyy-MM-dd}");
                return 0;
            }

            double referencePrice = 0;

            if (referenceSippRates.TryGetValue(intervalIndex, out var ratesForDates))
            {
                if (ratesForDates.TryGetValue(interval.StartDate, out referencePrice))
                {
                    Console.WriteLine($"  Found reference price for exact date: {referencePrice:F1}");
                }
                else
                {
                    var closestDate = ratesForDates.Keys
                        .OrderBy(d => Math.Abs((d - interval.StartDate).TotalDays))
                        .FirstOrDefault();

                    if (closestDate != default)
                    {
                        referencePrice = ratesForDates[closestDate];
                    }
                }
            }

            if (referencePrice <= 0)
            {
                Console.WriteLine($"  No reference price available for interval {intervalIndex}");
                return 0;
            }

            double factor = xdPrice / referencePrice;
            Console.WriteLine($"  Factor for {actualRentalLength}D: {xdPrice:F1} / {referencePrice:F1} = {factor:F3}");

            if (factor < 1)
            {
                Console.WriteLine($"  Adjusting factor from {factor:F3} to 1.000 (minimum allowed value)");
                factor = 1.0;
            }

            factor = Math.Ceiling(factor * 100) / 100;
            Console.WriteLine($"  Rounded factor to {factor:F3}");

            return factor;
        }

        private static Dictionary<int, Dictionary<string, double>> NormalizePricesForLinkedSipps(List<PriceInterval> intervals, List<HashSet<string>> linkedSippGroups)
        {
            var result = new Dictionary<int, Dictionary<string, double>>();

            for (int i = 0; i < intervals.Count; i++)
            {
                result[i] = new Dictionary<string, double>(intervals[i].PricesBySipp);
            }

            foreach (var linkedGroup in linkedSippGroups)
            {
                Console.WriteLine($"Processing linked group: {string.Join(", ", linkedGroup)}");

                for (int i = 0; i < intervals.Count; i++)
                {
                    double highestPrice = 0;
                    bool foundAny = false;

                    foreach (var sipp in linkedGroup)
                    {
                        if (intervals[i].PricesBySipp.TryGetValue(sipp, out double price))
                        {
                            highestPrice = Math.Max(highestPrice, price);
                            foundAny = true;
                        }
                    }

                    if (foundAny)
                    {
                        foreach (var sipp in linkedGroup)
                        {
                            if (intervals[i].PricesBySipp.ContainsKey(sipp))
                            {
                                if (result[i][sipp] != highestPrice)
                                {
                                    Console.WriteLine($"  Interval {i + 1}: Normalized {sipp} from {result[i][sipp]:F2} to {highestPrice:F2}");
                                    result[i][sipp] = highestPrice;
                                }
                            }
                        }
                    }
                }
            }

            return result;
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            string columnName = "";

            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                char letter = (char)('A' + remainder);
                columnName = letter + columnName;
                columnIndex = (columnIndex - 1) / 26;
            }

            return columnName;
        }
    }
}
