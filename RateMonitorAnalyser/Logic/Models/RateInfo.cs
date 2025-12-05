using System;

namespace RateMonitorAnalyser.Logic.Models
{
    public class RateInfo
    {
        public DateTime PickupDate { get; set; }

        public string Sipp { get; set; } = string.Empty;

        public double SuggestedAmount { get; set; }

        public string RuleDescription { get; set; } = string.Empty;

        public string Location { get; set; } = string.Empty;

        public int Lor { get; set; }

        public override string ToString()
        {
            return $"{PickupDate:yyyy-MM-dd} - {Sipp}: {SuggestedAmount:F2} ({RuleDescription})";
        }
    }
}
