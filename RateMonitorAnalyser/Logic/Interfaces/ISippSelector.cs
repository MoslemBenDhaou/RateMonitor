using System.Collections.Generic;

namespace RateMonitorAnalyser.Logic.Interfaces
{
    public interface ISippSelector
    {
        List<string> PromptForSippSelection(List<string> availableSipps);
    }
}
