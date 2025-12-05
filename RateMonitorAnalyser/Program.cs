using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Data;
using System.Globalization;

namespace RateMonitorAnalyser
{
    class Program
    {
        // Class to store rate information
        class RateInfo
        {
            public DateTime PickupDate { get; set; }
            public string Sipp { get; set; } = string.Empty;
            public double SuggestedAmount { get; set; }
            public string RuleDescription { get; set; } = string.Empty;
            public string Location { get; set; } = string.Empty;
            public int Lor { get; set; } // Length of rental
            
            public override string ToString()
            {
                return $"{PickupDate.ToString("yyyy-MM-dd")} - {Sipp}: {SuggestedAmount:F2} ({RuleDescription})";
            }
        }
        
        // Class to represent a price interval
        class PriceInterval
        {
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
            public Dictionary<string, double> PricesBySipp { get; set; } = new Dictionary<string, double>();
            public Dictionary<string, bool> IsFilledPrice { get; set; } = new Dictionary<string, bool>();
            public HashSet<string> ChangedSipps { get; set; } = new HashSet<string>(); // Sipp codes that caused this interval
            
            public override string ToString()
            {
                string changedSippsInfo = ChangedSipps.Any() 
                    ? $" - Changes: {string.Join(", ", ChangedSipps.OrderBy(s => s))}" 
                    : " - Initial interval";
                
                return $"{StartDate.ToString("yyyy-MM-dd")} to {EndDate.ToString("yyyy-MM-dd")} ({(EndDate - StartDate).TotalDays + 1} days){changedSippsInfo}";
            }
        }
        
        static void Main(string[] args)
        {
            // Set the EPPlus license for non-commercial use
            ExcelPackage.License.SetNonCommercialPersonal("RateMonitorAnalyser");
            
            Console.WriteLine("RateMonitor Analyser Starting...");
            
            // Define the source directory path
            string sourceDirectory = Path.Combine(Directory.GetCurrentDirectory(), "source");
            
            // Ensure the directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                Console.WriteLine("Creating source directory...");
                Directory.CreateDirectory(sourceDirectory);
                Console.WriteLine("Source directory created. Please place files in this directory and run the program again.");
                return;
            }
            
            // Define the pattern for the files we're looking for
            string filePattern = @"suggestion_report_\d+_\d{4}-\d{2}-\d{2}\.xlsx";
            Regex regex = new Regex(filePattern);
            
            // Get all files that match the pattern
            var matchingFiles = Directory.GetFiles(sourceDirectory)
                .Where(file => regex.IsMatch(Path.GetFileName(file)))
                .ToList();
            
            if (matchingFiles.Count == 0)
            {
                Console.WriteLine("No matching files found in the source directory.");
                Console.WriteLine("Expected file name format: suggestion_report_90468_2025-04-07.xlsx");
                return;
            }
            
            // Find the most recent file
            var mostRecentFile = matchingFiles
                .OrderByDescending(file => File.GetLastWriteTime(file))
                .First();
            
            Console.WriteLine($"Found most recent file: {Path.GetFileName(mostRecentFile)}");
            Console.WriteLine($"File path: {mostRecentFile}");
            Console.WriteLine($"Last modified: {File.GetLastWriteTime(mostRecentFile)}");
            
            // Extract the data
            var rateData = ExtractRateData(mostRecentFile);
            
            // Get all unique Sipp codes with data
            var allSipps = rateData
                .Where(r => r.SuggestedAmount > 0)
                .Select(r => r.Sipp)
                .Distinct()
                .OrderBy(s => s)
                .ToList();
            
            // Allow user to select which Sipp codes to include
            var selectedSipps = PromptForSippSelection(allSipps);
            
            // Analyze the data with selected Sipp codes
            var intervals = AnalyzePriceIntervals(rateData, selectedSipps);
            
            // Prompt for price adjustment factor
            Console.WriteLine("\nEnter price adjustment factor (e.g., 15 for 15% increase, -10 for 10% decrease, or 0 for no change):");
            string input = Console.ReadLine() ?? "0";
            
            double adjustmentFactor = 0;
            if (!double.TryParse(input, out adjustmentFactor))
            {
                Console.WriteLine("Invalid input. Using 0% adjustment (no change).");
                adjustmentFactor = 0;
            }
            
            // Convert percentage to multiplier (e.g., 15% -> 1.15, -10% -> 0.9)
            double priceMultiplier = 1 + (adjustmentFactor / 100);
            
            Console.WriteLine($"Applying {adjustmentFactor}% adjustment to all prices (multiplier: {priceMultiplier:F2})");
            
            // Export intervals and rates to CSV with price adjustment
            ExportIntervalsToCSV(intervals, selectedSipps, adjustmentFactor, priceMultiplier);
            
            // Prepare Excel template for export
            PrepareExcelTemplate(intervals, selectedSipps, adjustmentFactor, priceMultiplier);
            
            // Display intervals and allow user to select which one to view in detail
            DisplayIntervalsAndPromptForDetails(intervals);
            
            Console.WriteLine("Processing complete.");
        }
        
        static List<string> PromptForSippSelection(List<string> availableSipps)
        {
            Console.WriteLine("\nAvailable Sipp codes:");
            
            // Display all available Sipp codes with numbers
            for (int i = 0; i < availableSipps.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {availableSipps[i]}");
            }
            
            Console.WriteLine("\nEnter the numbers of the Sipp codes to include (comma-separated, or 'all' for all):");
            string input = Console.ReadLine() ?? "";
            
            if (input.Trim().ToLower() == "all")
            {
                Console.WriteLine("Including all Sipp codes in the analysis.");
                return availableSipps;
            }
            
            var selectedSipps = new List<string>();
            var selections = input.Split(',', StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var selection in selections)
            {
                if (int.TryParse(selection.Trim(), out int index) && index >= 1 && index <= availableSipps.Count)
                {
                    selectedSipps.Add(availableSipps[index - 1]);
                }
            }
            
            if (selectedSipps.Count == 0)
            {
                Console.WriteLine("No valid selections made. Including all Sipp codes.");
                return availableSipps;
            }
            
            Console.WriteLine($"Selected {selectedSipps.Count} Sipp codes: {string.Join(", ", selectedSipps)}");
            return selectedSipps;
        }
        
        static List<RateInfo> ExtractRateData(string filePath)
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
                    
                    // Get the dimensions of the worksheet
                    int rows = worksheet.Dimension?.Rows ?? 0;
                    int columns = worksheet.Dimension?.Columns ?? 0;
                    
                    // Find column indexes
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
                    
                    // Check if all required columns were found
                    if (sippColIndex == -1 || pickupDateColIndex == -1 || suggestedAmountColIndex == -1)
                    {
                        Console.WriteLine("Required columns not found in the Excel file.");
                        return rateData;
                    }
                    
                    // Process data rows - include ALL entries, even those with 0 values
                    for (int row = 2; row <= rows; row++)
                    {
                        // Extract cell values
                        var sipp = worksheet.Cells[row, sippColIndex].Value?.ToString();
                        var pickupDateValue = worksheet.Cells[row, pickupDateColIndex].Value;
                        var suggestedAmountValue = worksheet.Cells[row, suggestedAmountColIndex].Value;
                        var ruleDescription = worksheet.Cells[row, ruleDescriptionColIndex].Value?.ToString() ?? "";
                        var location = worksheet.Cells[row, locationColIndex].Value?.ToString() ?? "";
                        var lorValue = worksheet.Cells[row, lorColIndex].Value;
                        
                        // Skip rows with missing essential data
                        if (string.IsNullOrEmpty(sipp) || pickupDateValue == null)
                            continue;
                        
                        // Parse values
                        double suggestedAmount = 0;
                        if (suggestedAmountValue != null)
                        {
                            double.TryParse(suggestedAmountValue.ToString(), out suggestedAmount);
                            // Round to 0.1 (nearest 10 cents)
                            suggestedAmount = Math.Round(suggestedAmount * 10) / 10;
                        }
                        
                        // Convert Excel date to DateTime
                        DateTime pickupDate;
                        if (pickupDateValue is double excelDate)
                        {
                            pickupDate = DateTime.FromOADate(excelDate);
                        }
                        else if (!DateTime.TryParse(pickupDateValue.ToString(), out pickupDate))
                        {
                            continue;
                        }
                        
                        // Parse length of rental
                        int lor = 0;
                        if (lorValue != null)
                        {
                            int.TryParse(lorValue.ToString(), out lor);
                        }
                        
                        // Add to rate data list - include entries with 0 values
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
                
                // Sort by pickup date
                rateData = rateData.OrderBy(r => r.PickupDate).ToList();
                
                // Display date range
                if (rateData.Any())
                {
                    var firstDate = rateData.First().PickupDate;
                    var lastDate = rateData.Last().PickupDate;
                    Console.WriteLine($"Date range: {firstDate.ToString("yyyy-MM-dd")} to {lastDate.ToString("yyyy-MM-dd")} ({(lastDate - firstDate).TotalDays + 1} days)");
                    
                    // Count unique dates
                    var uniqueDates = rateData.Select(r => r.PickupDate.Date).Distinct().Count();
                    Console.WriteLine($"Number of unique dates: {uniqueDates}");
                    
                    // Count unique Sipp codes
                    var uniqueSipps = rateData.Select(r => r.Sipp).Distinct().Count();
                    Console.WriteLine($"Number of unique Sipp codes: {uniqueSipps}");
                    
                    // Count entries with 0 suggested amount
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
        
        static List<PriceInterval> AnalyzePriceIntervals(List<RateInfo> rateData, List<string> selectedSipps)
        {
            Console.WriteLine("\nAnalyzing price intervals...");
            
            var intervals = new List<PriceInterval>();
            
            if (!rateData.Any())
            {
                Console.WriteLine("No rate data to analyze.");
                return intervals;
            }
            
            try
            {
                // Get all unique dates
                var allDates = rateData.Select(r => r.PickupDate.Date).Distinct().OrderBy(d => d).ToList();
                
                Console.WriteLine($"Found {allDates.Count} unique dates and {selectedSipps.Count} selected Sipp codes.");
                
                // Create a dictionary to store raw prices by date and Sipp (including zeros)
                var rawPricesByDateAndSipp = new Dictionary<DateTime, Dictionary<string, double>>();
                
                // Initialize the dictionary for all dates
                foreach (var date in allDates)
                {
                    rawPricesByDateAndSipp[date] = new Dictionary<string, double>();
                }
                
                // Populate the dictionary with raw prices (including zeros)
                foreach (var rate in rateData)
                {
                    if (selectedSipps.Contains(rate.Sipp))
                    {
                        rawPricesByDateAndSipp[rate.PickupDate.Date][rate.Sipp] = rate.SuggestedAmount;
                    }
                }
                
                // Identify Sipp codes that have at least one non-zero value
                var sippsWithData = new HashSet<string>();
                foreach (var rate in rateData)
                {
                    if (selectedSipps.Contains(rate.Sipp) && rate.SuggestedAmount > 0)
                    {
                        sippsWithData.Add(rate.Sipp);
                    }
                }
                
                // Filter out Sipp codes with no data
                var validSipps = selectedSipps.Where(s => sippsWithData.Contains(s)).ToList();
                var ignoredSipps = selectedSipps.Where(s => !sippsWithData.Contains(s)).ToList();
                
                if (ignoredSipps.Any())
                {
                    Console.WriteLine($"Ignoring {ignoredSipps.Count} selected Sipp codes with no non-zero values: {string.Join(", ", ignoredSipps)}");
                    Console.WriteLine($"Analyzing {validSipps.Count} Sipp codes with at least one non-zero value");
                }
                
                // Create a dictionary to store filled prices by date and Sipp (with zeros replaced)
                var filledPricesByDateAndSipp = new Dictionary<DateTime, Dictionary<string, double>>();
                var isFilledPrice = new Dictionary<DateTime, Dictionary<string, bool>>();
                
                // Initialize the filled prices dictionary
                foreach (var date in allDates)
                {
                    filledPricesByDateAndSipp[date] = new Dictionary<string, double>();
                    isFilledPrice[date] = new Dictionary<string, bool>();
                }
                
                // Fill in missing prices using the closest past price or future price
                foreach (var sipp in validSipps)
                {
                    // Track the last non-zero price seen
                    double? lastNonZeroPrice = null;
                    
                    // First pass: fill in prices using past prices
                    foreach (var date in allDates)
                    {
                        if (rawPricesByDateAndSipp[date].TryGetValue(sipp, out double price) && price > 0)
                        {
                            // If we have a non-zero price, use it (already rounded during extraction)
                            filledPricesByDateAndSipp[date][sipp] = price;
                            isFilledPrice[date][sipp] = false; // Not filled
                            lastNonZeroPrice = price;
                        }
                        else if (lastNonZeroPrice.HasValue)
                        {
                            // If we have a zero or missing price but have seen a non-zero price in the past, use that
                            filledPricesByDateAndSipp[date][sipp] = lastNonZeroPrice.Value;
                            isFilledPrice[date][sipp] = true; // Filled
                        }
                        // Otherwise, leave it empty for now
                    }
                    
                    // Second pass: for dates that still have no price, fill in using future prices
                    // Find the first non-zero price in the future for each date that still has no price
                    for (int i = 0; i < allDates.Count; i++)
                    {
                        var date = allDates[i];
                        
                        // If this date doesn't have a price yet
                        if (!filledPricesByDateAndSipp[date].ContainsKey(sipp))
                        {
                            // Look for the first non-zero price in the future
                            double? futurePrice = null;
                            
                            for (int j = i + 1; j < allDates.Count; j++)
                            {
                                var futureDate = allDates[j];
                                if (rawPricesByDateAndSipp[futureDate].TryGetValue(sipp, out double price) && price > 0)
                                {
                                    futurePrice = price;
                                    break;
                                }
                            }
                            
                            // If we found a future price, use it
                            if (futurePrice.HasValue)
                            {
                                filledPricesByDateAndSipp[date][sipp] = futurePrice.Value;
                                isFilledPrice[date][sipp] = true; // Filled
                            }
                        }
                    }
                }
                
                // Create a dictionary to store normalized prices
                var normalizedPricesByDateAndSipp = new Dictionary<DateTime, Dictionary<string, double>>();
                
                // Initialize the normalized prices dictionary
                foreach (var date in allDates)
                {
                    normalizedPricesByDateAndSipp[date] = new Dictionary<string, double>();
                }
                
                // Process each Sipp code
                foreach (var sipp in validSipps)
                {
                    // First, collect all prices for this Sipp
                    var priceHistory = new Dictionary<DateTime, double>();
                    foreach (var date in allDates)
                    {
                        if (filledPricesByDateAndSipp[date].TryGetValue(sipp, out double price))
                        {
                            priceHistory[date] = price;
                        }
                    }
                    
                    if (priceHistory.Count < 2)
                    {
                        // Not enough data to normalize
                        foreach (var entry in priceHistory)
                        {
                            normalizedPricesByDateAndSipp[entry.Key][sipp] = entry.Value;
                        }
                        continue;
                    }
                    
                    // STEP 8: Eliminate fluctuations less than 1€
                    var datesList = priceHistory.Keys.OrderBy(d => d).ToList();
                    var normalizedPrices = new Dictionary<DateTime, double>();
                    
                    // Initialize with original prices
                    foreach (var date in datesList)
                    {
                        normalizedPrices[date] = priceHistory[date];
                    }
                    
                    // Iteratively fix small fluctuations until no more changes are needed
                    bool rateFixed;
                    do
                    {
                        rateFixed = false;
                        
                        for (int i = 0; i < datesList.Count - 1; i++)
                        {
                            var today = datesList[i];
                            var tomorrow = datesList[i + 1];
                            
                            double todayPrice = normalizedPrices[today];
                            double tomorrowPrice = normalizedPrices[tomorrow];
                            
                            // If difference is less than 1€, both dates should have the same value (the highest)
                            if (Math.Abs(todayPrice - tomorrowPrice) < 1.0)
                            {
                                double highestPrice = Math.Max(todayPrice, tomorrowPrice);
                                
                                // Only mark as fixed if we're actually changing a value
                                if (normalizedPrices[today] != highestPrice || normalizedPrices[tomorrow] != highestPrice)
                                {
                                    normalizedPrices[today] = highestPrice;
                                    normalizedPrices[tomorrow] = highestPrice;
                                    rateFixed = true;
                                    
                                    Console.WriteLine($"Fixed small fluctuation for {sipp}: {today.ToString("yyyy-MM-dd")} and {tomorrow.ToString("yyyy-MM-dd")} both set to {highestPrice:F2}");
                                }
                            }
                        }
                    } while (rateFixed); // Continue until no more fixes are needed
                    
                    // STEP 9: Detect and fix one-day fluctuations
                    // If today is different from both tomorrow and yesterday, today should match the highest of both
                    for (int i = 1; i < datesList.Count - 1; i++)
                    {
                        var yesterday = datesList[i - 1];
                        var today = datesList[i];
                        var tomorrow = datesList[i + 1];
                        
                        double yesterdayPrice = normalizedPrices[yesterday];
                        double todayPrice = normalizedPrices[today];
                        double tomorrowPrice = normalizedPrices[tomorrow];
                        
                        // Check if today's price is different from both yesterday and tomorrow by at least 1€
                        bool isDifferentFromYesterday = Math.Abs(todayPrice - yesterdayPrice) >= 1.0;
                        bool isDifferentFromTomorrow = Math.Abs(todayPrice - tomorrowPrice) >= 1.0;
                        
                        if (isDifferentFromYesterday && isDifferentFromTomorrow)
                        {
                            // One-day fluctuation detected - set to the highest of yesterday and tomorrow
                            double correctedPrice = Math.Max(yesterdayPrice, tomorrowPrice);
                            normalizedPrices[today] = correctedPrice;
                            
                            Console.WriteLine($"Fixed one-day fluctuation for {sipp} on {today.ToString("yyyy-MM-dd")}: {todayPrice:F2} -> {correctedPrice:F2}");
                        }
                    }
                    
                    // Store the final normalized prices
                    foreach (var entry in normalizedPrices)
                    {
                        normalizedPricesByDateAndSipp[entry.Key][sipp] = entry.Value;
                    }
                }
                
                // STEP 10: Find price change dates for each Sipp using the normalized prices
                var priceChangeDates = new Dictionary<DateTime, HashSet<string>>(); // Date -> Set of Sipps that changed price
                
                foreach (var sipp in validSipps)
                {
                    // Process the normalized price history day by day
                    var datesList = allDates.Where(d => normalizedPricesByDateAndSipp[d].ContainsKey(sipp)).OrderBy(d => d).ToList();
                    if (datesList.Count < 2)
                        continue;
                    
                    double? lastPrice = null;
                    
                    foreach (var date in datesList)
                    {
                        double currentPrice = normalizedPricesByDateAndSipp[date][sipp];
                        
                        if (lastPrice.HasValue && Math.Abs(lastPrice.Value - currentPrice) >= 1.0)
                        {
                            // Significant price change - start a new interval
                            if (!priceChangeDates.ContainsKey(date))
                            {
                                priceChangeDates[date] = new HashSet<string>();
                            }
                            priceChangeDates[date].Add(sipp);
                        }
                        
                        lastPrice = currentPrice;
                    }
                }
                
                // Add the first date to create a complete set of intervals
                if (!priceChangeDates.ContainsKey(allDates.First()))
                {
                    priceChangeDates[allDates.First()] = new HashSet<string>(); // No specific Sipp caused this, it's just the start
                }
                
                // Sort the price change dates
                var sortedChangeDates = priceChangeDates.Keys.OrderBy(d => d).ToList();
                
                // Create intervals
                for (int i = 0; i < sortedChangeDates.Count; i++)
                {
                    var startDate = sortedChangeDates[i];
                    var endDate = (i < sortedChangeDates.Count - 1) 
                        ? sortedChangeDates[i + 1].AddDays(-1) 
                        : allDates.Last();
                    
                    if (startDate <= endDate)
                    {
                        intervals.Add(new PriceInterval
                        {
                            StartDate = startDate,
                            EndDate = endDate,
                            ChangedSipps = priceChangeDates[startDate] // Store which Sipps caused this interval
                        });
                    }
                }
                
                // Populate prices for each interval using the normalized prices
                foreach (var interval in intervals)
                {
                    var sampleDate = interval.StartDate;
                    
                    foreach (var sipp in validSipps)
                    {
                        if (normalizedPricesByDateAndSipp[sampleDate].TryGetValue(sipp, out double price))
                        {
                            interval.PricesBySipp[sipp] = price;
                            
                            // Track if this was a filled price
                            if (isFilledPrice[sampleDate].TryGetValue(sipp, out bool filled))
                            {
                                interval.IsFilledPrice[sipp] = filled;
                            }
                        }
                    }
                }
                
                Console.WriteLine($"Identified {intervals.Count} price intervals.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing price intervals: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            
            return intervals;
        }
        
        static void DisplayIntervalsAndPromptForDetails(List<PriceInterval> intervals)
        {
            if (intervals.Count == 0)
            {
                Console.WriteLine("No intervals to display.");
                return;
            }
            
            // Display a list of all intervals
            Console.WriteLine("\nPrice Intervals:");
            Console.WriteLine("---------------");
            
            for (int i = 0; i < intervals.Count; i++)
            {
                var interval = intervals[i];
                Console.WriteLine($"{i + 1}. {interval}");
            }
            
            // Prompt user to select an interval to view in detail
            Console.WriteLine("\nEnter the interval number to view prices (or 0 to exit):");
            
            while (true)
            {
                string input = Console.ReadLine() ?? "";
                
                if (int.TryParse(input, out int selection))
                {
                    if (selection == 0)
                    {
                        break;
                    }
                    else if (selection >= 1 && selection <= intervals.Count)
                    {
                        DisplayIntervalPrices(intervals[selection - 1]);
                        Console.WriteLine("\nEnter another interval number to view prices (or 0 to exit):");
                    }
                    else
                    {
                        Console.WriteLine($"Please enter a number between 0 and {intervals.Count}:");
                    }
                }
                else
                {
                    Console.WriteLine("Please enter a valid number:");
                }
            }
        }
        
        static void DisplayIntervalPrices(PriceInterval interval)
        {
            Console.WriteLine($"\nInterval: {interval}");
            
            // Display which Sipp codes caused this interval (if any)
            if (interval.ChangedSipps.Any())
            {
                Console.WriteLine($"Caused by price changes in: {string.Join(", ", interval.ChangedSipps.OrderBy(s => s))}");
            }
            else
            {
                Console.WriteLine("This is the initial interval (no specific price changes)");
            }
            
            Console.WriteLine("Prices by Sipp:");
            
            foreach (var sipp in interval.PricesBySipp.Keys.OrderBy(s => s))
            {
                if (interval.PricesBySipp.TryGetValue(sipp, out double price))
                {
                    // Check if this was a filled price
                    bool isFilled = interval.IsFilledPrice.TryGetValue(sipp, out bool filled) && filled;
                    
                    string priceDisplay = isFilled ? $"{price:F2} (filled)" : $"{price:F2}";
                    Console.WriteLine($"  {sipp}: {priceDisplay}");
                }
                else
                {
                    Console.WriteLine($"  {sipp}: No data available");
                }
            }
        }
        
        static void AnalyzeRuleDescriptions(ExcelWorksheet worksheet)
        {
            try
            {
                // Find the column index for RuleDescription
                int ruleDescriptionColIndex = -1;
                int headerRow = 1;
                int columns = worksheet.Dimension?.Columns ?? 0;
                
                for (int col = 1; col <= columns; col++)
                {
                    var headerValue = worksheet.Cells[headerRow, col].Value?.ToString();
                    if (headerValue == "RuleDescription")
                    {
                        ruleDescriptionColIndex = col;
                        break;
                    }
                }
                
                if (ruleDescriptionColIndex == -1)
                {
                    Console.WriteLine("RuleDescription column not found.");
                    return;
                }
                
                // Count occurrences of each rule description
                var ruleCounts = new Dictionary<string, int>();
                int rows = worksheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rows; row++)
                {
                    var ruleDescription = worksheet.Cells[row, ruleDescriptionColIndex].Value?.ToString();
                    if (!string.IsNullOrEmpty(ruleDescription))
                    {
                        if (ruleCounts.ContainsKey(ruleDescription))
                        {
                            ruleCounts[ruleDescription]++;
                        }
                        else
                        {
                            ruleCounts[ruleDescription] = 1;
                        }
                    }
                }
                
                // Display rule description counts
                Console.WriteLine("\nRule Description Analysis:");
                foreach (var rule in ruleCounts.OrderByDescending(r => r.Value))
                {
                    Console.WriteLine($"  {rule.Key}: {rule.Value} occurrences");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing rule descriptions: {ex.Message}");
            }
        }
        
        static void AnalyzeLocations(ExcelWorksheet worksheet)
        {
            try
            {
                // Find the column index for Location
                int locationColIndex = -1;
                int headerRow = 1;
                int columns = worksheet.Dimension?.Columns ?? 0;
                
                for (int col = 1; col <= columns; col++)
                {
                    var headerValue = worksheet.Cells[headerRow, col].Value?.ToString();
                    if (headerValue == "Location")
                    {
                        locationColIndex = col;
                        break;
                    }
                }
                
                if (locationColIndex == -1)
                {
                    Console.WriteLine("Location column not found.");
                    return;
                }
                
                // Count occurrences of each location
                var locationCounts = new Dictionary<string, int>();
                int rows = worksheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rows; row++)
                {
                    var location = worksheet.Cells[row, locationColIndex].Value?.ToString();
                    if (!string.IsNullOrEmpty(location))
                    {
                        if (locationCounts.ContainsKey(location))
                        {
                            locationCounts[location]++;
                        }
                        else
                        {
                            locationCounts[location] = 1;
                        }
                    }
                }
                
                // Display location counts
                Console.WriteLine("\nLocation Analysis:");
                foreach (var location in locationCounts.OrderByDescending(l => l.Value).Take(10))
                {
                    Console.WriteLine($"  {location.Key}: {location.Value} occurrences");
                }
                
                if (locationCounts.Count > 10)
                {
                    Console.WriteLine($"  ... and {locationCounts.Count - 10} more locations");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing locations: {ex.Message}");
            }
        }
        
        static void AnalyzeCarTypes(ExcelWorksheet worksheet)
        {
            try
            {
                // Find the column index for Sipp (car type)
                int sippColIndex = -1;
                int headerRow = 1;
                int columns = worksheet.Dimension?.Columns ?? 0;
                
                for (int col = 1; col <= columns; col++)
                {
                    var headerValue = worksheet.Cells[headerRow, col].Value?.ToString();
                    if (headerValue == "Sipp")
                    {
                        sippColIndex = col;
                        break;
                    }
                }
                
                if (sippColIndex == -1)
                {
                    Console.WriteLine("Sipp column not found.");
                    return;
                }
                
                // Count occurrences of each car type
                var sippCounts = new Dictionary<string, int>();
                int rows = worksheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rows; row++)
                {
                    var sipp = worksheet.Cells[row, sippColIndex].Value?.ToString();
                    if (!string.IsNullOrEmpty(sipp))
                    {
                        if (sippCounts.ContainsKey(sipp))
                        {
                            sippCounts[sipp]++;
                        }
                        else
                        {
                            sippCounts[sipp] = 1;
                        }
                    }
                }
                
                // Display car type counts
                Console.WriteLine("\nCar Type (Sipp) Analysis:");
                foreach (var sipp in sippCounts.OrderByDescending(s => s.Value))
                {
                    Console.WriteLine($"  {sipp.Key}: {sipp.Value} occurrences");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing car types: {ex.Message}");
            }
        }
        
        static void ExportIntervalsToCSV(List<PriceInterval> intervals, List<string> sipps, double adjustmentFactor, double priceMultiplier)
        {
            if (intervals.Count == 0)
            {
                Console.WriteLine("No intervals to export.");
                return;
            }
            
            Console.WriteLine("\nExporting intervals and rates to CSV...");
            
            // Define linked Sipp categories
            var linkedSippGroups = new List<HashSet<string>>
            {
                new HashSet<string> { "EMMS", "PSMS" },
                new HashSet<string> { "PMMS", "PSAS" }
                // Add more linked groups as needed
            };
            
            Console.WriteLine("\nLinked Sipp categories (will have the same prices):");
            foreach (var group in linkedSippGroups)
            {
                Console.WriteLine($"  {string.Join(" and ", group)}");
            }
            
            // Create output directory if it doesn't exist
            string outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "output");
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
            
            // Create CSV file path with timestamp and adjustment factor
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string adjustmentString = adjustmentFactor >= 0 ? $"plus{adjustmentFactor}" : $"minus{Math.Abs(adjustmentFactor)}";
            string csvFilePath = Path.Combine(outputDirectory, $"price_intervals_{timestamp}_{adjustmentString}.csv");
            
            try
            {
                using (var writer = new StreamWriter(csvFilePath))
                {
                    // Write header row with interval dates
                    writer.Write("Sipp");
                    foreach (var interval in intervals)
                    {
                        writer.Write($",{interval.StartDate.ToString("yyyy-MM-dd")} to {interval.EndDate.ToString("yyyy-MM-dd")}");
                    }
                    writer.WriteLine();
                    
                    // Prepare a dictionary to store the adjusted prices for each Sipp and interval
                    var adjustedPrices = new Dictionary<string, Dictionary<int, double>>();
                    
                    // Initialize the adjusted prices dictionary
                    foreach (var sipp in sipps)
                    {
                        adjustedPrices[sipp] = new Dictionary<int, double>();
                    }
                    
                    // First pass: calculate adjusted prices for each Sipp and interval
                    for (int i = 0; i < intervals.Count; i++)
                    {
                        var interval = intervals[i];
                        
                        // Process each Sipp code
                        foreach (var sipp in sipps)
                        {
                            if (interval.PricesBySipp.TryGetValue(sipp, out double price))
                            {
                                // Apply the price adjustment factor
                                double adjustedPrice = price * priceMultiplier;
                                adjustedPrices[sipp][i] = adjustedPrice;
                            }
                        }
                    }
                    
                    // Second pass: handle linked Sipp categories
                    foreach (var linkedGroup in linkedSippGroups)
                    {
                        // Filter to only include Sipps that are in our selected list
                        var selectedLinkedSipps = linkedGroup.Where(s => sipps.Contains(s)).ToList();
                        
                        if (selectedLinkedSipps.Count <= 1)
                            continue; // Skip if there's only one or no Sipp from this group in our selection
                        
                        // For each interval, find the highest price among linked Sipps
                        for (int i = 0; i < intervals.Count; i++)
                        {
                            double highestPrice = 0;
                            
                            // Find the highest price in this linked group for this interval
                            foreach (var sipp in selectedLinkedSipps)
                            {
                                if (adjustedPrices[sipp].TryGetValue(i, out double price) && price > highestPrice)
                                {
                                    highestPrice = price;
                                }
                            }
                            
                            // If we found a highest price, set it for all Sipps in this linked group
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
                    
                    // Write the adjusted prices to the CSV
                    foreach (var sipp in sipps.OrderBy(s => s))
                    {
                        writer.Write(sipp);
                        
                        // Write price for each interval
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
                    
                    // Add information about the price adjustment and linked categories
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
        
        static void PrepareExcelTemplate(List<PriceInterval> intervals, List<string> sipps, double adjustmentFactor, double priceMultiplier)
        {
            Console.WriteLine("\nPreparing Excel template for export...");
            
            // Define the template path
            string exportDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Export");
            string templatePath = Path.Combine(exportDirectory, "TUN pricing template.xlsx");
            
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }
            
            // Create a new file name with today's date
            string todayDate = DateTime.Now.ToString("yyyy-MM-dd");
            string newFileName = $"TUN Pricing {todayDate}.xlsx";
            string newFilePath = Path.Combine(exportDirectory, newFileName);
            
            string referenceSipp = "ESMS";
            
            // Check if the reference Sipp is in our selected Sipps
            if (!sipps.Contains(referenceSipp))
            {
                Console.WriteLine($"Warning: Reference Sipp {referenceSipp} not found in selected Sipps. Using first available Sipp instead.");
                referenceSipp = sipps.FirstOrDefault() ?? "N/A";
            }
            
            Console.WriteLine($"Using {referenceSipp} as reference Sipp for row 8 pricing.");
            
            // Define linked Sipp categories
            var linkedSippGroups = new List<HashSet<string>>
            {
                //new HashSet<string> { "EMMS", "PSMS" },
                //new HashSet<string> { "PMMS", "PSAS" }
                // Add more linked groups as needed
            };
            
            Console.WriteLine("\nLinked Sipp categories (will have the same prices):");
            foreach (var group in linkedSippGroups)
            {
                Console.WriteLine($"  {string.Join(" and ", group)}");
            }
            
            try
            {
                // Copy the template to the new file
                File.Copy(templatePath, newFilePath, true);
                Console.WriteLine($"Created new Excel file: {newFileName}");
                
                // Find XD files for reserve length factors
                string sourceDirectory = Path.Combine(Directory.GetCurrentDirectory(), "source");
                var xdFiles = Directory.GetFiles(sourceDirectory, "*D_*.xlsx")
                    .Where(f => Regex.IsMatch(Path.GetFileName(f), @"^\d+D_.*\.xlsx$"))
                    .ToList();
                
                Console.WriteLine($"Found {xdFiles.Count} XD files for reserve length factors:");
                foreach (var file in xdFiles)
                {
                    Console.WriteLine($"  {Path.GetFileName(file)}");
                }
                
                // Extract XD data for reserve length factors
                var xdData = ExtractXDData(xdFiles, referenceSipp);
                
                // Open the new file and prepare it for editing
                using (var package = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    // Get the PRICING worksheet
                    var worksheet = package.Workbook.Worksheets["PRICING"];
                    
                    if (worksheet == null)
                    {
                        Console.WriteLine("PRICING worksheet not found in template.");
                        return;
                    }
                    
                    Console.WriteLine("Found PRICING worksheet in template.");
                    
                    // Apply linked Sipp logic to normalize prices
                    Console.WriteLine("Applying linked Sipp group logic to normalize prices...");
                    var normalizedPrices = NormalizePricesForLinkedSipps(intervals, linkedSippGroups);
                    
                    // Fill in the interval dates in the specified format
                    // Start dates in row 6, columns B, D, F, etc.
                    // End dates in row 7, columns B, D, F, etc.
                    // Reference Sipp 7-day rates in row 8, columns B, D, F, etc.
                    
                    // Sort intervals by start date to ensure chronological order
                    var sortedIntervals = intervals.OrderBy(i => i.StartDate).ToList();
                    
                    // Dictionary to store 7-day rates for the reference Sipp for each interval
                    var referenceSippRates = new Dictionary<int, Dictionary<DateTime, double>>();
                    
                    // First pass: Fill in dates and reference Sipp rates
                    for (int i = 0; i < sortedIntervals.Count; i++)
                    {
                        // Calculate the column letter (B, D, F, etc.)
                        // Column index starts at 1, B is column 2, D is column 4, etc.
                        int columnIndex = 2 + (i * 2); // B=2, D=4, F=6, etc.
                        string columnLetter = GetExcelColumnName(columnIndex);
                        
                        // Set start date in row 6
                        string startDateCell = $"{columnLetter}6";
                        worksheet.Cells[startDateCell].Value = sortedIntervals[i].StartDate;
                        worksheet.Cells[startDateCell].Style.Numberformat.Format = "yyyy-mm-dd";
                        
                        // Set end date in row 7
                        string endDateCell = $"{columnLetter}7";
                        worksheet.Cells[endDateCell].Value = sortedIntervals[i].EndDate;
                        worksheet.Cells[endDateCell].Style.Numberformat.Format = "yyyy-mm-dd";
                        
                        // Set 7-day rate for reference Sipp in row 8
                        string rateCell = $"{columnLetter}8";
                        
                        if (normalizedPrices.TryGetValue(i, out var pricesBySipp) && 
                            pricesBySipp.TryGetValue(referenceSipp, out double price))
                        {
                            // Calculate 7-day rate (price * 7)
                            double sevenDayRate = price * 7;
                            
                            // Apply the price adjustment factor
                            double adjustedSevenDayRate = sevenDayRate * priceMultiplier;
                            
                            worksheet.Cells[rateCell].Value = adjustedSevenDayRate;
                            worksheet.Cells[rateCell].Style.Numberformat.Format = "#,##0.00";
                            
                            // Store the reference rate for later percentage calculations
                            if (!referenceSippRates.ContainsKey(i))
                            {
                                referenceSippRates[i] = new Dictionary<DateTime, double>();
                            }
                            
                            // Store the daily rate for each date in the interval
                            for (DateTime date = sortedIntervals[i].StartDate; date <= sortedIntervals[i].EndDate; date = date.AddDays(1))
                            {
                                referenceSippRates[i][date] = price;
                            }
                            
                            Console.WriteLine($"Filled interval {i+1}: {columnLetter}6 = {sortedIntervals[i].StartDate.ToString("yyyy-MM-dd")}, " +
                                             $"{columnLetter}7 = {sortedIntervals[i].EndDate.ToString("yyyy-MM-dd")}, " +
                                             $"{columnLetter}8 = {adjustedSevenDayRate:F2} ({referenceSipp} 7-day rate)");
                        }
                        else
                        {
                            worksheet.Cells[rateCell].Value = "N/A";
                            Console.WriteLine($"Filled interval {i+1}: {columnLetter}6 = {sortedIntervals[i].StartDate.ToString("yyyy-MM-dd")}, " +
                                             $"{columnLetter}7 = {sortedIntervals[i].EndDate.ToString("yyyy-MM-dd")}, " +
                                             $"{columnLetter}8 = N/A (No price data for {referenceSipp})");
                        }
                    }
                    
                    Console.WriteLine("All intervals filled in PRICING sheet.");
                    
                    // Second pass: Fill in percentage differences in columns C, E, G, etc. based on existing Sipp codes in column A
                    // Start from row 9 and read existing Sipp codes
                    int nextRow = 9;
                    bool hasMoreRows = true;
                    
                    while (hasMoreRows)
                    {
                        // Read the Sipp code from column A
                        var sippCell = worksheet.Cells[$"A{nextRow}"];
                        string? sipp = sippCell.Value?.ToString();
                        
                        // If cell is empty or we've reached the end of the data, break the loop
                        if (string.IsNullOrWhiteSpace(sipp))
                        {
                            hasMoreRows = false;
                            continue;
                        }
                        
                        Console.WriteLine($"Found Sipp code in A{nextRow}: {sipp}");
                        
                        // For each interval, calculate and set the percentage difference
                        for (int i = 0; i < sortedIntervals.Count; i++)
                        {
                            // Calculate the column letter for the rate (B, D, F, etc.)
                            int rateColumnIndex = 2 + (i * 2);
                            string? rateColumnLetter = GetExcelColumnName(rateColumnIndex);
                            
                            // Calculate the column letter for the percentage (C, E, G, etc.)
                            int percentageColumnIndex = 3 + (i * 2);
                            string? percentageColumnLetter = GetExcelColumnName(percentageColumnIndex);
                            
                            if (rateColumnLetter != null && percentageColumnLetter != null)
                            {
                                // Set the rate and percentage difference
                                string rateCell = $"{rateColumnLetter}{nextRow}";
                                string percentageCell = $"{percentageColumnLetter}{nextRow}";
                                
                                // Check if we have both the reference rate and this Sipp's rate
                                if (referenceSippRates.TryGetValue(i, out var ratesForDates) && 
                                    normalizedPrices.TryGetValue(i, out var pricesBySipp))
                                {
                                    // Get the reference rate (using the first date in the interval)
                                    double referenceRate = ratesForDates[sortedIntervals[i].StartDate] * 7;
                                    
                                    // Apply special multipliers for specific Sipp codes
                                    double sippPrice;
                                    double sippMultiplier = 1.0;
                                    
                                    if (sipp == "PMAS")
                                    {
                                        // PMAS should be 3 times the price of ESMS
                                        sippMultiplier = 3.0;
                                        sippPrice = referenceRate * sippMultiplier;
                                        Console.WriteLine($"Applied special multiplier for {sipp}: {sippMultiplier}x");
                                    }
                                    else if (sipp == "SMAS")
                                    {
                                        // SMAS should be 4 times the price of ESMS
                                        sippMultiplier = 4.0;
                                        sippPrice = referenceRate * sippMultiplier;
                                        Console.WriteLine($"Applied special multiplier for {sipp}: {sippMultiplier}x");
                                    }
                                    else if (pricesBySipp.TryGetValue(sipp, out double price))
                                    {
                                        // Use the actual price from the data
                                        sippPrice = price * 7;
                                    }
                                    else
                                    {
                                        // No price available for this Sipp
                                        worksheet.Cells[rateCell].Value = "0";
                                        worksheet.Cells[percentageCell].Value = 0;
                                        worksheet.Cells[percentageCell].Style.Numberformat.Format = "0.0%";
                                        Console.WriteLine($"No price available for {sipp} in interval {i+1}");
                                        continue;
                                    }
                                    
                                    // Calculate 7-day rate for this Sipp
                                    double sippSevenDayRate = sippPrice;
                                    
                                    // Calculate percentage difference using original prices
                                    double percentageDiff = 0;
                                    if (referenceRate > 0)
                                    {
                                        percentageDiff = ((sippSevenDayRate - referenceRate) / referenceRate) * 100;
                                    }
                                    
                                    // For display in the Excel sheet, apply the adjustment to the reference rate and Sipp rate
                                    double adjustedReferenceRate = referenceRate * priceMultiplier;
                                    double adjustedSippSevenDayRate = sippSevenDayRate * priceMultiplier;
                                    
                                    // Set the actual price in the rate cell
                                    worksheet.Cells[rateCell].Value = adjustedSippSevenDayRate;
                                    worksheet.Cells[rateCell].Style.Numberformat.Format = "#,##0.00";
                                    
                                    // Set the percentage in the percentage cell
                                    worksheet.Cells[percentageCell].Value = percentageDiff / 100; // Excel needs decimal for percentage format
                                    worksheet.Cells[percentageCell].Style.Numberformat.Format = "0.0%";
                                    
                                    Console.WriteLine($"Set {sipp} price for interval {i+1}: {adjustedSippSevenDayRate:F2} in cell {rateCell}");
                                    Console.WriteLine($"Set {sipp} percentage difference to {referenceSipp} for interval {i+1}: {percentageDiff:F1}% in cell {percentageCell}");
                                    Console.WriteLine($"  Original prices: {sipp}={sippSevenDayRate:F2}, {referenceSipp}={referenceRate:F2}");
                                    Console.WriteLine($"  Adjusted prices: {sipp}={adjustedSippSevenDayRate:F2}, {referenceSipp}={adjustedReferenceRate:F2}");
                                }
                                else
                                {
                                    worksheet.Cells[rateCell].Value = "N/A";
                                    worksheet.Cells[percentageCell].Value = 0;
                                    worksheet.Cells[percentageCell].Style.Numberformat.Format = "0.0%";
                                    Console.WriteLine($"Could not calculate price and percentage for {sipp} in interval {i+1} - missing data");
                                }
                            }
                        }
                        
                        // Move to the next row
                        nextRow++;
                        
                        // Safety check to avoid infinite loops (stop after 100 rows)
                        if (nextRow > 100)
                        {
                            Console.WriteLine("Reached maximum row limit (100). Stopping.");
                            break;
                        }
                    }
                    
                    // Save the workbook to ensure the PRICING sheet changes are persisted
                    package.Save();
                    Console.WriteLine($"Saved Excel template with PRICING sheet data to {templatePath}");
                    
                    // Now fill the RESERVE LENGTH sheet
                    var reserveSheet = package.Workbook.Worksheets["RESERVE LENGTH"];
                    if (reserveSheet == null)
                    {
                        Console.WriteLine("RESERVE LENGTH worksheet not found in template.");
                    }
                    else
                    {
                        Console.WriteLine("Found RESERVE LENGTH worksheet. Filling reserve length factors...");
                        
                        // Map rental lengths to column letters
                        var rentalLengthToColumn = new Dictionary<int, string?>
                        {
                            { 1, "B" },
                            { 2, "C" },
                            { 3, "D" },
                            { 4, "E" },
                            { 5, "F" },
                            { 6, "G" }
                        };
                        
                        // Find the next available row for each rental length
                        var nextRowForRentalLength = new Dictionary<int, int>();
                        
                        // Initialize with starting row (likely row 2 or 3, but we'll search to be sure)
                        foreach (var rentalLength in rentalLengthToColumn.Keys)
                        {
                            string? column = rentalLengthToColumn[rentalLength];
                            if (column == null) continue;
                            
                            // Start at row 2 for all rental lengths
                            nextRowForRentalLength[rentalLength] = 2;
                            Console.WriteLine($"Starting row for {rentalLength}D (column {column}): 2");
                        }
                        
                        // Define the row spacing between intervals
                        const int rowSpacingBetweenIntervals = 13; // To get 2, 15, 28, 41, etc.
                        
                        // Fill in the factors for each interval and rental length
                        foreach (var interval in sortedIntervals)
                        {
                            Console.WriteLine($"Processing interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}");
                            
                            // For each rental length (1-6 days)
                            for (int rentalLength = 1; rentalLength <= 6; rentalLength++)
                            {
                                if (rentalLengthToColumn.TryGetValue(rentalLength, out string? column) && column != null)
                                {
                                    // Find the factor for this rental length
                                    double factor = CalculateReserveLengthFactor(
                                        interval, 
                                        rentalLength, 
                                        referenceSipp, 
                                        xdData, 
                                        referenceSippRates);
                                    
                                    // Ensure factor is never less than 1
                                    if (factor > 0 && factor < 1)
                                    {
                                        Console.WriteLine($"  Adjusting factor from {factor:F3} to 1.000 (minimum allowed value)");
                                        factor = 1.0;
                                    }
                                    
                                    // Round the factor up to the nearest 0.01
                                    factor = Math.Ceiling(factor * 100) / 100;
                                    Console.WriteLine($"  Rounded factor to {factor:F3}");
                                    
                                    if (factor > 0)
                                    {
                                        // Get the next available row for this rental length
                                        if (nextRowForRentalLength.TryGetValue(rentalLength, out int rowNum))
                                        {
                                            // Set the factor in the cell
                                            string factorCell = $"{column}{rowNum}";
                                            reserveSheet.Cells[factorCell].Value = factor;
                                            reserveSheet.Cells[factorCell].Style.Numberformat.Format = "0.00";
                                            
                                            Console.WriteLine($"Set {rentalLength}D factor for interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}: {factor:F3} in cell {factorCell}");
                                            
                                            // Increment the row for next time
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
                    
                    // Save the workbook to ensure the PRICING sheet changes are persisted
                    package.Save();
                    Console.WriteLine($"Saved Excel template with PRICING sheet data to {templatePath}");
                    
                    // Save the changes to the workbook
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
        
        // Helper method to extract data from XD files
        static Dictionary<int, Dictionary<DateTime, double>> ExtractXDData(List<string> xdFiles, string referenceSipp)
        {
            var result = new Dictionary<int, Dictionary<DateTime, double>>();
            
            foreach (var file in xdFiles)
            {
                // Extract the rental length from the filename (e.g., "1D_2025-04-07.xlsx" -> 1)
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
                                // Initialize dictionary for this rental length
                                result[rentalLength] = new Dictionary<DateTime, double>();
                                
                                // Find the columns we need
                                int sippColumn = -1;
                                int pickUpDateColumn = -1;
                                int suggestedAmountColumn = -1;
                                
                                // Find the header row and column indices
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
                                    // Process each row
                                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        string sipp = worksheet.Cells[row, sippColumn].Value?.ToString() ?? string.Empty;
                                        
                                        if (sipp == referenceSipp)
                                        {
                                            DateTime pickUpDate;
                                            
                                            // Handle different types of date values
                                            if (worksheet.Cells[row, pickUpDateColumn].Value is DateTime dateValue)
                                            {
                                                pickUpDate = dateValue;
                                            }
                                            else if (worksheet.Cells[row, pickUpDateColumn].Value is double numericDate)
                                            {
                                                // Convert Excel numeric date to DateTime
                                                pickUpDate = DateTime.FromOADate(numericDate);
                                            }
                                            else if (worksheet.Cells[row, pickUpDateColumn].Value is int intDate)
                                            {
                                                // Convert integer date to DateTime
                                                pickUpDate = DateTime.FromOADate(intDate);
                                            }
                                            else if (worksheet.Cells[row, pickUpDateColumn].Value is string dateStr && 
                                                     DateTime.TryParse(dateStr, out DateTime parsedDate))
                                            {
                                                pickUpDate = parsedDate;
                                            }
                                            else
                                            {
                                                // Skip this row if we can't parse the date
                                                continue;
                                            }
                                            
                                            // Try to extract the suggested amount
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
                                                // Skip this row if we can't parse the amount
                                                continue;
                                            }
                                            
                                            // Round the suggested amount to the nearest 0.1
                                            suggestedAmount = Math.Round(suggestedAmount * 10) / 10;
                                            
                                            // Store the daily price for this date
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
        
        // Helper method to calculate reserve length factor
        static double CalculateReserveLengthFactor(
            PriceInterval interval, 
            int rentalLength, 
            string referenceSipp, 
            Dictionary<int, Dictionary<DateTime, double>> xdData,
            Dictionary<int, Dictionary<DateTime, double>> referenceSippRates)
        {
            Console.WriteLine($"Calculating factor for interval {interval.StartDate:yyyy-MM-dd} to {interval.EndDate:yyyy-MM-dd}, rental length {rentalLength}D");
            
            // If we don't have data for this rental length, try to find the next higher rental length
            int actualRentalLength = rentalLength;
            while (actualRentalLength <= 6)
            {
                if (xdData.ContainsKey(actualRentalLength))
                    break;
                
                Console.WriteLine($"  No data for {actualRentalLength}D, trying {actualRentalLength + 1}D");
                actualRentalLength++;
            }
            
            // If we still don't have data, return 0
            if (actualRentalLength > 6 || !xdData.ContainsKey(actualRentalLength))
            {
                Console.WriteLine($"  No data available for any rental length >= {rentalLength}D");
                return 0;
            }
            
            if (actualRentalLength != rentalLength)
            {
                Console.WriteLine($"  Using {actualRentalLength}D data for {rentalLength}D factor");
            }
            
            // Find the XD price for the interval start date
            if (!xdData[actualRentalLength].TryGetValue(interval.StartDate, out double xdPrice))
            {
                // If we don't have an exact match, try to find the closest date
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
                    return 0; // No data available
                }
            }
            else
            {
                Console.WriteLine($"  Found exact price for {actualRentalLength}D on {interval.StartDate:yyyy-MM-dd}: {xdPrice:F1}");
            }
            
            // Find the interval index in referenceSippRates
            int intervalIndex = -1;
            
            // First, try to find the exact interval by start date
            foreach (var entry in referenceSippRates)
            {
                // Check if this dictionary contains the start date of our interval
                if (entry.Value.ContainsKey(interval.StartDate))
                {
                    intervalIndex = entry.Key;
                    Console.WriteLine($"  Found interval index by start date: {intervalIndex}");
                    break;
                }
            }
            
            // If not found, try to find the interval that contains the start date
            if (intervalIndex == -1)
            {
                Console.WriteLine($"  Could not find exact interval by start date, trying to find containing interval");
                
                // Try each interval
                for (int i = 0; i < referenceSippRates.Count; i++)
                {
                    if (referenceSippRates.TryGetValue(i, out var intervalRates))
                    {
                        // Get the dates in this interval
                        var dates = intervalRates.Keys.OrderBy(d => d).ToList();
                        
                        if (dates.Count > 0)
                        {
                            // Check if our interval start date falls within this interval's dates
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
                return 0; // Could not find interval index
            }
            
            // Get the reference price for the interval
            double referencePrice = 0;
            
            if (referenceSippRates.TryGetValue(intervalIndex, out var ratesForDates))
            {
                // First try to get the exact date
                if (ratesForDates.TryGetValue(interval.StartDate, out referencePrice))
                {
                    Console.WriteLine($"  Found reference price for exact date: {referencePrice:F1}");
                }
                else
                {
                    // Try to find the closest date
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
                return 0; // No reference price available
            }
            
            // Calculate the factor
            double factor = xdPrice / referencePrice;
            Console.WriteLine($"  Factor for {actualRentalLength}D: {xdPrice:F1} / {referencePrice:F1} = {factor:F3}");
            
            // Ensure factor is never less than 1
            if (factor < 1)
            {
                Console.WriteLine($"  Adjusting factor from {factor:F3} to 1.000 (minimum allowed value)");
                factor = 1.0;
            }
            
            // Round the factor up to the nearest 0.01
            factor = Math.Ceiling(factor * 100) / 100;
            Console.WriteLine($"  Rounded factor to {factor:F3}");
            
            return factor;
        }
        
        // Helper method to normalize prices for linked Sipps
        static Dictionary<int, Dictionary<string, double>> NormalizePricesForLinkedSipps(List<PriceInterval> intervals, List<HashSet<string>> linkedSippGroups)
        {
            var result = new Dictionary<int, Dictionary<string, double>>();
            
            // Initialize result with original prices
            for (int i = 0; i < intervals.Count; i++)
            {
                result[i] = new Dictionary<string, double>(intervals[i].PricesBySipp);
            }
            
            // Process each linked group
            foreach (var linkedGroup in linkedSippGroups)
            {
                Console.WriteLine($"Processing linked group: {string.Join(", ", linkedGroup)}");
                
                // For each interval
                for (int i = 0; i < intervals.Count; i++)
                {
                    // Find the highest price among linked Sipps for this interval
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
                    
                    // If we found any prices, set all linked Sipps to the highest price
                    if (foundAny)
                    {
                        foreach (var sipp in linkedGroup)
                        {
                            if (intervals[i].PricesBySipp.ContainsKey(sipp))
                            {
                                if (result[i][sipp] != highestPrice)
                                {
                                    Console.WriteLine($"  Interval {i+1}: Normalized {sipp} from {result[i][sipp]:F2} to {highestPrice:F2}");
                                    result[i][sipp] = highestPrice;
                                }
                            }
                        }
                    }
                }
            }
            
            return result;
        }
        
        // Helper method to convert column index to Excel column name (A, B, C, ... Z, AA, AB, etc.)
        static string GetExcelColumnName(int columnIndex)
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
