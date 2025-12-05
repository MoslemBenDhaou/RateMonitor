using System;
using System.Collections.Generic;
using System.Linq;

namespace RateMonitorAnalyser.Logic.Models
{
    public class PriceInterval
    {
        public DateTime StartDate { get; set; }

        public DateTime EndDate { get; set; }

        public Dictionary<string, double> PricesBySipp { get; set; } = new Dictionary<string, double>();

        public Dictionary<string, bool> IsFilledPrice { get; set; } = new Dictionary<string, bool>();

        public HashSet<string> ChangedSipps { get; set; } = new HashSet<string>();

        public override string ToString()
        {
            string changedSippsInfo = ChangedSipps.Any()
                ? $" - Changes: {string.Join(", ", ChangedSipps.OrderBy(s => s))}"
                : " - Initial interval";

            return $"{StartDate:yyyy-MM-dd} to {EndDate:yyyy-MM-dd} ({(EndDate - StartDate).TotalDays + 1} days){changedSippsInfo}";
        }
    }
}
