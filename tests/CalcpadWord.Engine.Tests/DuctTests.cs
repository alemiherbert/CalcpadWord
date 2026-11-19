using System.Globalization;

namespace CalcpadWord.Engine.Tests;

/// <summary>
/// The Dokimos duct sizing script (templates/dokimos/Template source/duct, linked as
/// Fixtures/Duct.cpd), calculated as the add-in does in Word.
/// </summary>
public class DuctTests
{
    private static readonly TemplateScript Duct = new("Duct.cpd");

    /// <summary>How the duct is sized, by the numbers of the drop-down list.</summary>
    private const int ByFriction = 1, ByVelocity = 2, ByDiameter = 3;

    /// <summary>McQuay DuctSizer 6.4's roughness: its friction factors fit Colebrook with 0.0007 ft.</summary>
    private const double DuctSizerRoughness = 0.21336;

    private static Calculation Size(double q, int by, double x, double ε = 0.09, double t = 20, double z = 0, double a = 500, params string[] probes) =>
        Duct.Calculate(new()
        {
            ["t"] = Text(t),
            ["z"] = Text(z),
            ["ε"] = Text(ε),
            ["q"] = Text(q),
            ["P_s"] = Text(by),
            ["x_s"] = Text(x),
            ["a"] = Text(a),
        }, decimals: 6, probes: probes);

    private static string Text(double x) => x.ToString("0.######", CultureInfo.InvariantCulture);

    [Fact]
    public void Has_input_fields_for_the_air_the_duct_and_the_sizing() =>
        Assert.Equal(["t", "z", "ε", "q", "P_s", "x_s", "a"], Duct.InputNames);

    [Fact]
    public void The_default_duct_calculates_without_errors()
    {
        var c = Duct.Calculate();

        Assert.False(c.HasErrors);
        Assert.Contains("Friction loss", c.LastBlock);
    }

    [Fact]
    public void Air_at_20_degrees_at_sea_level()
    {
        var c = Size(500, ByFriction, 1, probes: ["ρ", "μ*10^5"]); // μ alone shows as 1.813×10⁻⁵

        Assert.Equal(1.204, c.Probed("ρ"), 0.001);
        Assert.Equal(1.813, c.Probed("μ*10^5"), 0.002);
    }

    /// <summary>
    /// McQuay DuctSizer 6.4 (20 °C air at STP, metric). It uses slightly different air
    /// (1.2014 kg/m³, 1.786·10⁻⁵ Pa·s) and stops its iteration near the target (0.995 for 1 Pa/m),
    /// so diameters agree within 0.3 % and friction within 1 %.
    /// </summary>
    [Theory]
    [InlineData(500, ByFriction, 1, 346.2, 0.02029, 1)]
    [InlineData(100, ByFriction, 1, 189.2, 0.02469, 1)]
    [InlineData(2000, ByFriction, 0.8, 611.2, 0.01741, 0.8)]
    [InlineData(1000, ByVelocity, 6, 460.7, 0.01876, 0.882)]
    public void Agrees_with_McQuay_DuctSizer(double q, int by, double x, double d, double f, double loss)
    {
        var c = Size(q, by, x, ε: DuctSizerRoughness, probes: ["D", "f", "Δp_L"]);

        Assert.False(c.HasErrors);
        Assert.Equal(d, c.Probed("D"), d * 0.003);
        Assert.Equal(f, c.Probed("f"), f * 0.01);
        Assert.Equal(loss, c.Probed("Δp_L"), loss * 0.01);
    }

    [Fact]
    public void A_given_diameter_gives_its_velocity_and_friction()
    {
        var c = Size(500, ByDiameter, 400, probes: ["V", "Δp_L"]);

        Assert.Equal(500 / 1000.0 / (Math.PI * 0.4 * 0.4 / 4), c.Probed("V"), 0.001);
        Assert.InRange(c.Probed("Δp_L"), 0.4, 0.6); // smaller than at 1 Pa/m: D 400 > 332 mm
    }

    /// <summary>Huebscher, as in DuctSizer: 500 × 200 mm is 337.03 mm round.</summary>
    [Fact]
    public void Equivalent_diameter_of_a_rectangular_duct() =>
        Assert.Equal(337.03, Size(500, ByFriction, 1, probes: ["D_eq(500; 200)"]).Probed("D_eq(500; 200)"), 0.01);

    [Fact]
    public void The_rectangular_side_is_rounded_up_so_the_friction_is_no_more()
    {
        var c = Size(500, ByFriction, 1, a: 500, probes: ["b_0", "b", "Δp_L,r"]);

        Assert.Equal(25 * Math.Ceiling(c.Probed("b_0") / 25), c.Probed("b"));
        Assert.True(c.Probed("b") >= c.Probed("b_0"));
        Assert.True(c.Probed("Δp_L,r") <= 1);
    }

    [Fact]
    public void Altitude_lowers_the_density_and_the_size()
    {
        var sea = Size(500, ByFriction, 1, z: 0, probes: ["D"]).Probed("D");
        var kampala = Size(500, ByFriction, 1, z: 1190, probes: ["D"]).Probed("D");

        Assert.True(kampala < sea);
    }

    [Theory]
    [InlineData(0, 1, "Choose how to size the duct")]
    [InlineData(ByFriction, 0, "must be above 0")]
    public void Bad_inputs_say_why(int by, double x, string message)
    {
        var c = Size(500, by, x);

        Assert.Contains(message, c.Blocks);
        Assert.False(c.HasErrors);
    }
}
