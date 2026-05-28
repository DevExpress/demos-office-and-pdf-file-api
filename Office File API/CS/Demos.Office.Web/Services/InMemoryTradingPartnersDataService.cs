namespace BlazorDemo.OfficeFileApi; 
sealed class InMemoryTradingPartnersDataService {
    readonly List<TradingPartner> tradingPartners = new() {
        new TradingPartner("Australia", 2277.9, 815.4, 1462.5, -0.0894, -0.0279, -0.0012, "July-2016"),
        new TradingPartner("Belgium", 2808.9, 1826.2, 982.7, 0.1923, 0.2291, 0.0013, "May-2016"),
        new TradingPartner("Brazil", 3423.2, 2571.4, 851.8, -0.0464, 0.1883, -0.004, "May-2016"),
        new TradingPartner("Canada", 26453.1, 28343.1, -1890, 0.1887, 0.0994, -0.0046, "May-2016"),
        new TradingPartner("China", 8786.5, 36646.2, -27859.7, -0.1478, 0.0518, 0.0014, "May-2016"),
        new TradingPartner("France", 2642.3, 3563.7, -921.4, 0.0882, 0.063, 0, "May-2016"),
        new TradingPartner("Germany", 4052.5, 9888.3, -5835.8, 0.0662, 0.052, 0.0004, "May-2016"),
        new TradingPartner("Hong Kong", 3385.4, 413, 2972.4, -0.2327, 0.0034, -0.0026, "May-2016"),
        new TradingPartner("India", 1935.8, 4201.5, -2265.7, -0.1334, 0.5152, 0.032, "May-2016"),
        new TradingPartner("Italy", 1581.1, 3292.1,  -1711,0.193,0.1067,0.0004,"May-2016"),
        new TradingPartner("Japan", 5785.2, 11190.4, -5405.2, -0.0392, -0.0529, -0.0007, "May-2016"),
        new TradingPartner("Mexico", 19240.5, 24540.6, -5300.1, 0.1745, 0.2101, 0.0036, "May-2016"),
        new TradingPartner("Netherlands", 3442.5, 1538.4, 1904.1, -0.1083, -0.0372, -0.0016, "May-2016"),
        new TradingPartner("Saudi Arabia", 1787.6, 4493, -2705.4, -0.0678, 0.2276, 0.0055, "May-2016"),
        new TradingPartner("Singapore", 2678.7, 1467.2, 1211.5, -0.017, -0.0005, -0.0004, "May-2016"),
        new TradingPartner("South Korea", 3211.5, 5669.9, -2458.4, -0.07, 0.2336, 0.0115, "May-2016"),
        new TradingPartner("Switzerland", 2150.8, 3249.1, -1098.3, 0.167, 0.5306, 0.0293, "May-2016")
    };
    public List<TradingPartner> TradingPartners => tradingPartners;
}
