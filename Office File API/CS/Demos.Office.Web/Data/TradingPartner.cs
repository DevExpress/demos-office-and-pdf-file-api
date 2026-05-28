namespace BlazorDemo.OfficeFileApi;

public class TradingPartner {
    public TradingPartner(string country, double exports, double imports, double balance, double exports1YChg, double imports1YChg, double balance1YChg, string asOf) {
        Country = country;
        Exports = exports;
        Imports = imports;
        Balance = balance;
        Exports1YChg = exports1YChg;
        Imports1YChg = imports1YChg;
        Balance1YChg = balance1YChg;
        AsOf = asOf;
    }
    public string Country { get; set; }
    public double Exports { get; set; }
    public double Imports { get; set; }
    public double Balance { get; set; }
    public double Exports1YChg { get; set; }
    public double Imports1YChg { get; set; }
    public double Balance1YChg { get; set; }
    public string AsOf { get; set; }
}
