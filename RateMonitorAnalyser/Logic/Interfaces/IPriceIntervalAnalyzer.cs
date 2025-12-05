using System.Collections.Generic;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Interfaces
{
    public interface IPriceIntervalAnalyzer
    {
        List<PriceInterval> AnalyzePriceIntervals(List<RateInfo> rateData, List<string> selectedSipps);
    }
}
