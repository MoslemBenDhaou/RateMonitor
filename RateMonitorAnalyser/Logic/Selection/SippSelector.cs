using System;
using System.Collections.Generic;

namespace RateMonitorAnalyser.Logic.Selection
{
    public class SippSelector : Interfaces.ISippSelector
    {
        public List<string> PromptForSippSelection(List<string> availableSipps)
        {
            Console.WriteLine("\nAvailable Sipp codes:");

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
    }
}
