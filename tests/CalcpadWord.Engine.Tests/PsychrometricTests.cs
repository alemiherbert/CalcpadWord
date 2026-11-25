using System.Globalization;

namespace CalcpadWord.Engine.Tests;

/// <summary>
/// The psychrometric chart's script (templates/Template source/psychrometric,
/// linked as Fixtures/Psychrometric.cpd), calculated as the add-in does in Word.
/// </summary>
public class PsychrometricTests
{
    private static readonly TemplateScript Chart = new("Psychrometric.cpd");

    /// <summary>The known properties, by the numbers of the drop-down lists.</summary>
    private const int DryBulb = 1, Humidity = 2, WetBulb = 3, DewPoint = 4, Ratio = 5, Enthalpy = 6, Volume = 7;

    private static Calculation Air(double z, int p1, double x1, int p2, double x2, params string[] probes) =>
        Chart.Calculate(new()
        {
            ["z"] = Text(z),
            ["P_1"] = Text(p1),
            ["x_1"] = Text(x1),
            ["P_2"] = Text(p2),
            ["x_2"] = Text(x2),
        }, decimals: 6, probes: probes);

    private static string Text(double x) => x.ToString("0.######", CultureInfo.InvariantCulture);

    [Fact]
    public void Has_input_fields_for_the_site_and_two_known_properties() =>
        Assert.Equal(["z", "P_1", "x_1", "P_2", "x_2"], Chart.InputNames);

    [Fact]
    public void The_default_air_calculates_without_errors()
    {
        var c = Chart.Calculate();

        Assert.False(c.HasErrors);
        Assert.Contains("Saturation, 100 %", c.LastBlock); // the chart's legend
    }

    /// <summary>ASHRAE Handbook – Fundamentals (SI), chapter 1, Table 1: standard atmosphere.</summary>
    [Theory]
    [InlineData(0, 101.325)]
    [InlineData(1000, 89.876)]
    [InlineData(1500, 84.556)]
    public void Atmospheric_pressure_is_the_standard_atmosphere(double z, double p) =>
        Assert.Equal(p, Air(z, DryBulb, 25, Humidity, 50, "p").Probed("p"), 0.005);

    /// <summary>ASHRAE, chapter 1, Table 3: saturation pressure, over ice below 0 °C.</summary>
    [Fact]
    public void Saturation_pressure_matches_the_ashrae_table()
    {
        var c = Air(0, DryBulb, 25, Humidity, 50, "p_sat(20)", "p_sat(0)", "p_sat(-10)", "p_sat(50)");

        // Within 0.05 %: newer tables (IAPWS) differ from Hyland and Wexler by about 0.02 % at 50 °C.
        Assert.Equal(2.3389, c.Probed("p_sat(20)"), 0.0012);
        Assert.Equal(0.61121, c.Probed("p_sat(0)"), 0.0003);
        Assert.Equal(0.2599, c.Probed("p_sat(-10)"), 0.00013);
        Assert.Equal(12.352, c.Probed("p_sat(50)"), 0.006);
    }

    /// <summary>Air at 25 °C and 50 % at sea level, as read from ASHRAE's chart.</summary>
    [Fact]
    public void Air_at_25_degrees_and_50_percent_at_sea_level()
    {
        var c = Air(0, DryBulb, 25, Humidity, 50, "W_g", "h", "t_wb", "t_d", "v");

        Assert.InRange(c.Probed("W_g"), 9.87, 9.89);
        Assert.InRange(c.Probed("h"), 50.2, 50.4);
        Assert.InRange(c.Probed("t_wb"), 17.85, 17.95);
        Assert.InRange(c.Probed("t_d"), 13.8, 13.9);
        Assert.InRange(c.Probed("v"), 0.8575, 0.8585);
    }

    private static readonly Lazy<double[]> Reference = new(() =>
    {
        var c = Air(0, DryBulb, 25, Humidity, 50, "t", "φ", "t_wb", "t_d", "W_g", "h", "v");
        return [0, c.Probed("t"), c.Probed("φ"), c.Probed("t_wb"), c.Probed("t_d"), c.Probed("W_g"), c.Probed("h"), c.Probed("v")];
    });

    /// <summary>The air at sea level, 25 °C and 50 %, found again from any two of its properties that fix it.</summary>
    [Theory]
    [InlineData(DryBulb, WetBulb)]
    [InlineData(DryBulb, DewPoint)]
    [InlineData(DryBulb, Ratio)]
    [InlineData(DryBulb, Enthalpy)]
    [InlineData(DryBulb, Volume)]
    [InlineData(Humidity, DryBulb)]
    [InlineData(Humidity, WetBulb)]
    [InlineData(Humidity, DewPoint)]
    [InlineData(Humidity, Ratio)]
    [InlineData(Humidity, Enthalpy)]
    [InlineData(Humidity, Volume)]
    [InlineData(WetBulb, DewPoint)]
    [InlineData(WetBulb, Ratio)]
    [InlineData(WetBulb, Volume)]
    [InlineData(DewPoint, Enthalpy)]
    [InlineData(DewPoint, Volume)]
    [InlineData(Ratio, Enthalpy)]
    [InlineData(Ratio, Volume)]
    [InlineData(Enthalpy, Volume)]
    public void Any_two_properties_that_fix_the_air_give_the_same_air(int first, int second)
    {
        var x = Reference.Value;
        var c = Air(0, first, x[first], second, x[second], "t", "W_g");

        Assert.False(c.HasErrors);
        Assert.Equal(25, c.Probed("t"), 1);
        Assert.Equal(x[Ratio], c.Probed("W_g"), 2);
    }

    [Theory]
    [InlineData(DewPoint, 10, Ratio, 8, "fix only the moisture")]
    [InlineData(WetBulb, 18, Enthalpy, 50, "almost the same lines")]
    [InlineData(DryBulb, 25, DryBulb, 30, "two different")]
    [InlineData(DryBulb, 25, Humidity, 150, "No air has these values")]
    [InlineData(DryBulb, 25, WetBulb, 30, "No air has these values")]
    public void Pairs_that_do_not_fix_the_air_say_why(int p1, double x1, int p2, double x2, string message)
    {
        var c = Air(0, p1, x1, p2, x2);

        Assert.Contains(message, c.Blocks);
        Assert.Contains("No chart", c.LastBlock);
        Assert.False(c.HasErrors); // inputs that don't fit are not a calculation that went wrong
    }
}
