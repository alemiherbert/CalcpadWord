using System.Globalization;

namespace CalcpadWord.Engine.Tests;

/// <summary>
/// The soil stresses and settlement script (templates/Template source/soil, linked
/// as Fixtures/Soil.cpd), after Calcpad's example "Soil stress and deformations".
/// </summary>
public class SoilTests
{
    private static readonly TemplateScript Soil = new("Soil.cpd");

    private static Calculation Load(double p1, double p2, double h = 6, params string[] probes) =>
        Soil.Calculate(new()
        {
            ["p_1"] = Text(p1),
            ["p_2"] = Text(p2),
            ["H"] = Text(h),
        }, decimals: 6, probes: probes);

    private static string Text(double x) => x.ToString("0.######", CultureInfo.InvariantCulture);

    [Fact]
    public void Has_input_fields_for_the_soil_the_area_and_the_load() =>
        Assert.Equal(["E_0", "ν", "a", "b", "H", "p_1", "p_2"], Soil.InputNames);

    [Fact]
    public void The_default_load_calculates_without_errors()
    {
        var c = Soil.Calculate();

        Assert.False(c.HasErrors);
        Assert.Contains("Under the edge", c.LastBlock);
    }

    /// <summary>
    /// Calcpad's example (its load, 100 kPa·(1 + x/a), is 50 to 150 kPa) integrates σ_z over the depth
    /// from 0.01·H: 9.96 mm at the centre. The template's closed form over z gives the same from there;
    /// from just below the surface it gives about 3 % more, the strain of the top 6 cm.
    /// </summary>
    [Fact]
    public void Agrees_with_Calcpads_example()
    {
        const string example = "(1 - ν^2)/E_0*$Integral{σ_z(0m; 0m; z) @ z = 0.01*H : H}|mm";
        const string closedForm = "(1 - ν^2)/E_0*$Area{$Area{p(x_p; y_p)*(K(-x_p; -y_p; H) - K(-x_p; -y_p; 0.01*H)) @ x_p = -a/2 : a/2} @ y_p = -b/2 : b/2}|mm";
        var c = Load(50, 150, probes: [example, closedForm, "d(0m; 0m)"]);

        Assert.Equal(9.96, c.Probed(example), 0.01);
        Assert.Equal(c.Probed(example), c.Probed(closedForm), 0.01);
        Assert.Equal(10.23, c.Probed("d(0m; 0m)"), 0.02);
    }

    /// <summary>
    /// A flexible rectangle under a uniform load on a deep elastic half-space (Timoshenko and Goodier;
    /// Steinbrenner): the corner of an L × B rectangle settles p·B·(1 − ν²)/E·I, with
    /// I = (m·ln((1 + √(1 + m²))/m) + ln(m + √(1 + m²)))/π and m = L/B. The centre of 2 × 3 m is four
    /// corners of 1 × 1.5 m: I = 0.6788 (1.36 for the whole width, as tabulated), 12.35 mm.
    /// </summary>
    [Fact]
    public void Agrees_with_the_closed_form_for_a_uniform_load_on_a_deep_half_space()
    {
        double m = 1.5, root = Math.Sqrt(1 + m * m);
        var i = (m * Math.Log((1 + root) / m) + Math.Log(m + root)) / Math.PI;
        var expected = 4 * 100 * 1.0 * (1 - 0.3 * 0.3) / 20000 * i * 1000; // mm

        var c = Load(100, 100, h: 1000, probes: ["d(0m; 0m)"]);

        Assert.Equal(12.35, expected, 0.01);
        Assert.Equal(expected, c.Probed("d(0m; 0m)"), expected * 0.005);
    }

    /// <summary>Far below, the load acts as a point load P: σ_z = 3P/(2π·z²) (0.1 % less, for its spread).</summary>
    [Fact]
    public void Far_below_the_stress_is_that_of_a_point_load()
    {
        var pointLoad = 3 * (100 * 2 * 3) / (2 * Math.PI * 50 * 50); // kPa
        var c = Load(100, 100, probes: ["σ_z(0m; 0m; 50m)"]);

        Assert.Equal(pointLoad, c.Probed("σ_z(0m; 0m; 50m)"), pointLoad * 0.003);
    }

    [Fact]
    public void The_heavier_edge_settles_more()
    {
        var c = Load(50, 150, probes: ["d(-a/2; 0m)", "d(a/2; 0m)", "d_max", "d(0m; 0m)"]);

        Assert.True(c.Probed("d(a/2; 0m)") > c.Probed("d(-a/2; 0m)"));
        Assert.True(c.Probed("d_max") >= c.Probed("d(0m; 0m)"));
    }
}
