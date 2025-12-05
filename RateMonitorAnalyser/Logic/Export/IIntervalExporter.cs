using System.Collections.Generic;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Export
{
    public interface IIntervalExporter
    {
        void Export(List<PriceInterval> intervals, List<string> sipps, double adjustmentFactor, double priceMultiplier, RateMonitorOptions options);
    }
}
