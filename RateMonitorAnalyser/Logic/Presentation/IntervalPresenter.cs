using System;
using System.Collections.Generic;
using System.Linq;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Presentation
{
    public class IntervalPresenter
    {
        public void DisplayIntervalsAndPromptForDetails(List<PriceInterval> intervals)
        {
            if (intervals.Count == 0)
            {
                Console.WriteLine("No intervals to display.");
                return;
            }

            Console.WriteLine("\nPrice Intervals:");
            Console.WriteLine("---------------");

            for (int i = 0; i < intervals.Count; i++)
            {
                var interval = intervals[i];
                Console.WriteLine($"{i + 1}. {interval}");
            }

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

        public void DisplayIntervalPrices(PriceInterval interval)
        {
            Console.WriteLine($"\nInterval: {interval}");

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
    }
}
