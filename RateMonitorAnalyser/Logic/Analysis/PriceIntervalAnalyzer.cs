using System;
using System.Collections.Generic;
using System.Linq;
using RateMonitorAnalyser.Logic.Interfaces;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Analysis
{
    public class PriceIntervalAnalyzer : IPriceIntervalAnalyzer
    {
        public List<PriceInterval> AnalyzePriceIntervals(List<RateInfo> rateData, List<string> selectedSipps)
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
                var allDates = rateData.Select(r => r.PickupDate.Date).Distinct().OrderBy(d => d).ToList();

                Console.WriteLine($"Found {allDates.Count} unique dates and {selectedSipps.Count} selected Sipp codes.");

                var rawPricesByDateAndSipp = new Dictionary<DateTime, Dictionary<string, double>>();

                foreach (var date in allDates)
                {
                    rawPricesByDateAndSipp[date] = new Dictionary<string, double>();
                }

                foreach (var rate in rateData)
                {
                    if (selectedSipps.Contains(rate.Sipp))
                    {
                        rawPricesByDateAndSipp[rate.PickupDate.Date][rate.Sipp] = rate.SuggestedAmount;
                    }
                }

                var sippsWithData = new HashSet<string>();
                foreach (var rate in rateData)
                {
                    if (selectedSipps.Contains(rate.Sipp) && rate.SuggestedAmount > 0)
                    {
                        sippsWithData.Add(rate.Sipp);
                    }
                }

                var validSipps = selectedSipps.Where(s => sippsWithData.Contains(s)).ToList();
                var ignoredSipps = selectedSipps.Where(s => !sippsWithData.Contains(s)).ToList();

                if (ignoredSipps.Any())
                {
                    Console.WriteLine($"Ignoring {ignoredSipps.Count} selected Sipp codes with no non-zero values: {string.Join(", ", ignoredSipps)}");
                    Console.WriteLine($"Analyzing {validSipps.Count} Sipp codes with at least one non-zero value");
                }

                var filledPricesByDateAndSipp = new Dictionary<DateTime, Dictionary<string, double>>();
                var isFilledPrice = new Dictionary<DateTime, Dictionary<string, bool>>();

                foreach (var date in allDates)
                {
                    filledPricesByDateAndSipp[date] = new Dictionary<string, double>();
                    isFilledPrice[date] = new Dictionary<string, bool>();
                }

                foreach (var sipp in validSipps)
                {
                    double? lastNonZeroPrice = null;

                    foreach (var date in allDates)
                    {
                        if (rawPricesByDateAndSipp[date].TryGetValue(sipp, out double price) && price > 0)
                        {
                            filledPricesByDateAndSipp[date][sipp] = price;
                            isFilledPrice[date][sipp] = false;
                            lastNonZeroPrice = price;
                        }
                        else if (lastNonZeroPrice.HasValue)
                        {
                            filledPricesByDateAndSipp[date][sipp] = lastNonZeroPrice.Value;
                            isFilledPrice[date][sipp] = true;
                        }
                    }

                    for (int i = 0; i < allDates.Count; i++)
                    {
                        var date = allDates[i];

                        if (!filledPricesByDateAndSipp[date].ContainsKey(sipp))
                        {
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

                            if (futurePrice.HasValue)
                            {
                                filledPricesByDateAndSipp[date][sipp] = futurePrice.Value;
                                isFilledPrice[date][sipp] = true;
                            }
                        }
                    }
                }

                var normalizedPricesByDateAndSipp = new Dictionary<DateTime, Dictionary<string, double>>();

                foreach (var date in allDates)
                {
                    normalizedPricesByDateAndSipp[date] = new Dictionary<string, double>();
                }

                foreach (var sipp in validSipps)
                {
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
                        foreach (var entry in priceHistory)
                        {
                            normalizedPricesByDateAndSipp[entry.Key][sipp] = entry.Value;
                        }
                        continue;
                    }

                    var datesList = priceHistory.Keys.OrderBy(d => d).ToList();
                    var normalizedPrices = new Dictionary<DateTime, double>();

                    foreach (var date in datesList)
                    {
                        normalizedPrices[date] = priceHistory[date];
                    }

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

                            if (Math.Abs(todayPrice - tomorrowPrice) < 1.0)
                            {
                                double highestPrice = Math.Max(todayPrice, tomorrowPrice);

                                if (normalizedPrices[today] != highestPrice || normalizedPrices[tomorrow] != highestPrice)
                                {
                                    normalizedPrices[today] = highestPrice;
                                    normalizedPrices[tomorrow] = highestPrice;
                                    rateFixed = true;

                                    Console.WriteLine($"Fixed small fluctuation for {sipp}: {today:yyyy-MM-dd} and {tomorrow:yyyy-MM-dd} both set to {highestPrice:F2}");
                                }
                            }
                        }
                    } while (rateFixed);

                    for (int i = 1; i < datesList.Count - 1; i++)
                    {
                        var yesterday = datesList[i - 1];
                        var today = datesList[i];
                        var tomorrow = datesList[i + 1];

                        double yesterdayPrice = normalizedPrices[yesterday];
                        double todayPrice = normalizedPrices[today];
                        double tomorrowPrice = normalizedPrices[tomorrow];

                        bool isDifferentFromYesterday = Math.Abs(todayPrice - yesterdayPrice) >= 1.0;
                        bool isDifferentFromTomorrow = Math.Abs(todayPrice - tomorrowPrice) >= 1.0;

                        if (isDifferentFromYesterday && isDifferentFromTomorrow)
                        {
                            double correctedPrice = Math.Max(yesterdayPrice, tomorrowPrice);
                            normalizedPrices[today] = correctedPrice;

                            Console.WriteLine($"Fixed one-day fluctuation for {sipp} on {today:yyyy-MM-dd}: {todayPrice:F2} -> {correctedPrice:F2}");
                        }
                    }

                    foreach (var entry in normalizedPrices)
                    {
                        normalizedPricesByDateAndSipp[entry.Key][sipp] = entry.Value;
                    }
                }

                var priceChangeDates = new Dictionary<DateTime, HashSet<string>>();

                foreach (var sipp in validSipps)
                {
                    var datesList = allDates.Where(d => normalizedPricesByDateAndSipp[d].ContainsKey(sipp)).OrderBy(d => d).ToList();
                    if (datesList.Count < 2)
                        continue;

                    double? lastPrice = null;

                    foreach (var date in datesList)
                    {
                        double currentPrice = normalizedPricesByDateAndSipp[date][sipp];

                        if (lastPrice.HasValue && Math.Abs(lastPrice.Value - currentPrice) >= 1.0)
                        {
                            if (!priceChangeDates.ContainsKey(date))
                            {
                                priceChangeDates[date] = new HashSet<string>();
                            }
                            priceChangeDates[date].Add(sipp);
                        }

                        lastPrice = currentPrice;
                    }
                }

                if (!priceChangeDates.ContainsKey(allDates.First()))
                {
                    priceChangeDates[allDates.First()] = new HashSet<string>();
                }

                var sortedChangeDates = priceChangeDates.Keys.OrderBy(d => d).ToList();

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
                            ChangedSipps = priceChangeDates[startDate]
                        });
                    }
                }

                foreach (var interval in intervals)
                {
                    var sampleDate = interval.StartDate;

                    foreach (var sipp in validSipps)
                    {
                        if (normalizedPricesByDateAndSipp[sampleDate].TryGetValue(sipp, out double price))
                        {
                            interval.PricesBySipp[sipp] = price;

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
    }
}
