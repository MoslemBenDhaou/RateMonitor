using System.Collections.Generic;
using RateMonitorAnalyser.Logic.Models;

namespace RateMonitorAnalyser.Logic.Interfaces
{
    public interface IRateDataExtractor
    {
        List<RateInfo> ExtractRateData(string filePath);
    }
}
