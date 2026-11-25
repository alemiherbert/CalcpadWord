using System.Globalization;

namespace CalcpadWord.Engine.Tests;

/// <summary>
/// The seepage script (templates/Template source/seepage, linked as
/// Fixtures/Seepage.cpd): steady confined seepage under a sheet pile wall, by finite differences.
/// It is checked against the exact solution, against the node method of Knappett's spreadsheet
/// Seepage_CSM8 (Craig's Soil Mechanics, 8th edition), solved independently here, and against
/// the result its manual quotes.
/// </summary>
public class SeepageTests
{
    private static readonly TemplateScript Seepage = new("Seepage.cpd");

    private static Calculation Run(Dictionary<string, double> values, params string[] probes) =>
        Seepage.Calculate(values.ToDictionary(v => v.Key, v => v.Value.ToString("0.########", CultureInfo.InvariantCulture)),
            decimals: 6, probes: probes);

    [Fact]
    public void Has_input_fields_for_the_soil_the_geometry_and_the_water() =>
        Assert.Equal(["k_1", "t_1", "k_2", "r_k", "γ_sat", "T", "d", "e_x", "L_1", "L_2", "Δ", "h_1", "h_2"], Seepage.InputNames);

    [Fact]
    public void The_defaults_calculate_without_errors_and_draw_the_flow_net()
    {
        var c = Seepage.Calculate();

        Assert.False(c.HasErrors, c.Errors);
        Assert.Contains("PASS", c.Blocks);
        Assert.Contains("Upstream face", c.LastBlock);
    }

    private static double HalfDepth(double delta) =>
        Run(new() { ["T"] = 4, ["d"] = 2, ["t_1"] = 4, ["L_1"] = 20, ["L_2"] = 20, ["Δ"] = delta, ["h_1"] = 1, ["h_2"] = 0 }, "q_n").Probed("q_n");

    /// <summary>
    /// A wall to half the depth of the layer: by symmetry, q = k·ΔH/2 exactly (Polubarinova-Kochina:
    /// q/(k·ΔH) = K(cos α)/(2·K(sin α)), α = π·d/(2T)); a section five depths wide each side. The
    /// flow round the toe is singular, so the error of the grid falls with the cell size, not its
    /// square: it halves with each halving, and two grids extrapolate to the exact flow.
    /// </summary>
    [Fact]
    public void Converges_to_the_exact_solution()
    {
        double coarse = HalfDepth(0.25), fine = HalfDepth(0.125);

        Assert.Equal(0.5, ExactShapeFactor(4, 2), 1e-12);
        Assert.InRange(0.5 - fine, 0.4 * (0.5 - coarse), 0.6 * (0.5 - coarse));
        Assert.Equal(0.5, 2 * fine - coarse, 0.001);
        Assert.Equal(0.5, fine, 0.01);
    }

    /// <summary>
    /// The node method of Seepage_CSM8 errs as much the other way (its toe lies half a cell higher):
    /// on the same grid, the mean of the two is within 0.2 % of the exact flow.
    /// </summary>
    [Fact]
    public void Brackets_the_exact_solution_with_the_spreadsheets_node_method()
    {
        var cells = HalfDepth(0.25);
        var nodes = NodeMethod(t: 4, d: 2, l1: 20, l2: 20, delta: 0.25, h1: 1, h2: 0);

        Assert.True(cells < 0.5 && nodes > 0.5, $"cells {cells}, nodes {nodes}");
        Assert.Equal(0.5, (cells + nodes) / 2, 0.001);
    }

    /// <summary>
    /// Craig's sheet pile, 6 m into 8.5 m of soil, 4.5 m and 0.5 m of water, 8 m each side: the node
    /// method on the 0.5 m grid of Seepage_CSM8 gives the flow its manual quotes, q = 1.4·k. The
    /// template and the node method, each extrapolated from the 0.5 m and 0.25 m grids, agree.
    /// (Over a layer of any width the exact flow is 1.488·k.)
    /// </summary>
    [Fact]
    public void Extrapolates_to_the_flow_of_the_spreadsheets_method()
    {
        double Cells(double delta) => Run(new() { ["T"] = 8.5, ["t_1"] = 8.5, ["L_1"] = 8, ["L_2"] = 8, ["Δ"] = delta }, "q_n").Probed("q_n");
        double Nodes(double delta) => NodeMethod(t: 8.5, d: 6, l1: 8, l2: 8, delta: delta, h1: 4.5, h2: 0.5);

        var nodesCoarse = Nodes(0.5);
        Assert.Equal(1.4, nodesCoarse, 0.02);
        var cells = 2 * Cells(0.25) - Cells(0.5);
        var nodes = 2 * Nodes(0.25) - nodesCoarse;
        Assert.Equal(nodes, cells, nodes * 0.005);
        Assert.Equal(1.488, ExactShapeFactor(8.6, 6) * 4, 0.001);
    }

    [Fact]
    public void What_flows_in_upstream_flows_under_the_toe()
    {
        var c = Run(new() { ["e_x"] = 1.4, ["t_1"] = 4, ["k_2"] = 2e-6, ["r_k"] = 2 }, "q_n", "q_in");

        Assert.Equal(c.Probed("q_n"), c.Probed("q_in"), 1e-6);
    }

    /// <summary>
    /// With k_h = 4·k_v the section is the isotropic one with its widths halved and k = 2·k_v
    /// (the transformed section): twice the flow of that section, to the accuracy of the grid.
    /// </summary>
    [Fact]
    public void Anisotropy_matches_the_transformed_section()
    {
        var anisotropic = Run(new() { ["r_k"] = 4, ["L_1"] = 24, ["L_2"] = 24 }, "q_n").Probed("q_n");
        var transformed = Run(new() { ["L_1"] = 12, ["L_2"] = 12, ["Δ"] = 0.1 }, "q_n").Probed("q_n");

        Assert.Equal(2 * transformed, anisotropic, 2 * transformed * 0.03);
    }

    [Fact]
    public void A_less_permeable_lower_layer_passes_less_flow()
    {
        var one = Seepage.Calculate(decimals: 6, probes: ["q_n"]).Probed("q_n");
        var two = Run(new() { ["t_1"] = 7, ["k_2"] = 1e-6 }, "q_n").Probed("q_n");

        Assert.True(two < 0.8 * one, $"{two} should be well below {one}");
    }

    [Fact]
    public void A_deeper_wall_passes_less_flow_and_is_safer_against_heave()
    {
        var shallow = Run(new() { ["d"] = 4 }, "q_n", "η_hv", "η_pip");
        var deep = Run(new() { ["d"] = 7 }, "q_n", "η_hv", "η_pip");

        Assert.True(deep.Probed("q_n") < shallow.Probed("q_n"));
        Assert.True(deep.Probed("η_hv") < shallow.Probed("η_hv"));
        Assert.True(deep.Probed("η_pip") < shallow.Probed("η_pip"));
    }

    /// <summary>A cofferdam: 3 m excavated in front of the wall, the water in it at the excavation level.</summary>
    [Fact]
    public void An_excavation_in_front_of_the_wall_calculates()
    {
        var c = Run(new() { ["e_x"] = 3, ["h_1"] = 2, ["h_2"] = 0 }, "q_n", "q_in", "i_e", "P_w");

        Assert.False(c.HasErrors);
        Assert.Equal(c.Probed("q_n"), c.Probed("q_in"), 1e-6);
        Assert.True(c.Probed("i_e") > 0);
        Assert.True(c.Probed("P_w") > 0);
    }

    /// <summary>
    /// Under the toe the two faces meet one pressure, and above the ground the water is hydrostatic:
    /// on the upstream face γ_w·h_1 at the ground.
    /// </summary>
    [Fact]
    public void The_pressures_on_the_wall_are_hydrostatic_above_the_ground_and_meet_at_the_toe()
    {
        var c = Seepage.Calculate(decimals: 6, probes: ["u_u.2", "u_u.(n_u)", "u_w.(n_w)", "P_w", "a_w"]);

        Assert.Equal(9.81 * 4.5, c.Probed("u_u.2"), 1e-6);
        Assert.Equal(c.Probed("u_u.(n_u)"), c.Probed("u_w.(n_w)"), 1e-6);
        Assert.InRange(c.Probed("a_w"), 0, 10.5);
    }

    // ---- Reference: the node method of Seepage_CSM8, solved by successive over-relaxation -----

    /// <summary>
    /// q/k for a wall of zero thickness on a node column, as in Seepage_CSM8: each node the mean of
    /// its neighbours, a no-flow boundary mirrored (its inner neighbour counts twice), a node beside
    /// the wall mirrored across it, and the toe the mean of its left, right and lower neighbours and
    /// of the two nodes just above it, one each side of the wall. The heads of the top row are those
    /// of the water. The flow is what enters the upstream ground (half at the edge and at the wall,
    /// where a node has half a cell).
    /// </summary>
    private static double NodeMethod(double t, double d, double l1, double l2, double delta, double h1, double h2)
    {
        int nx = (int)Math.Round((l1 + l2) / delta), nz = (int)Math.Round(t / delta);
        int cw = (int)Math.Round(l1 / delta), rd = (int)Math.Round(d / delta);
        var h = new double[nz + 1, nx + 1];      // heads; on the wall column above the toe, the left side
        var right = new double[rd + 1];          // the right side of the wall column above the toe
        for (var c = 0; c <= nx; c++)
            for (var r = 0; r <= nz; r++)
                h[r, c] = c < cw ? h1 : h2;
        h[0, cw] = h1;
        Array.Fill(right, h2);

        double Up(int r, int c, bool rightSide) => c == cw && r < rd ? (rightSide ? right[r] : h[r, c]) : h[r, c];

        for (var sweep = 0; sweep < 100000; sweep++)
        {
            var change = 0.0;
            for (var r = 1; r <= nz; r++)
            {
                for (var c = 0; c <= nx; c++)
                {
                    double down(bool side) => r < nz ? Up(r + 1, c, side) : Up(r - 1, c, side); // mirrored at the base
                    double value;
                    if (c == cw && r < rd)
                    {
                        // Beside the wall: the wall face mirrors the neighbour on its own side.
                        var left = (2 * h[r, c - 1] + Up(r - 1, c, false) + (r + 1 < rd ? Up(r + 1, c, false) : h[r + 1, c])) / 4;
                        var rightValue = (2 * h[r, c + 1] + Up(r - 1, c, true) + (r + 1 < rd ? Up(r + 1, c, true) : h[r + 1, c])) / 4;
                        change = Math.Max(change, Math.Max(Math.Abs(left - h[r, c]), Math.Abs(rightValue - right[r])));
                        h[r, c] += 1.9 * (left - h[r, c]);
                        right[r] += 1.9 * (rightValue - right[r]);
                        continue;
                    }
                    // Just right of the wall, the neighbour on the wall column is its right side.
                    var west = c == cw + 1 && r < rd ? right[r] : c > 0 ? h[r, c - 1] : h[r, c + 1];
                    var east = c < nx ? h[r, c + 1] : h[r, c - 1];
                    var north = c == cw && r == rd ? (h[r - 1, c] + right[r - 1]) / 2 : h[r - 1, c];
                    value = (west + east + north + down(false)) / 4;
                    change = Math.Max(change, Math.Abs(value - h[r, c]));
                    h[r, c] += 1.9 * (value - h[r, c]);
                }
            }
            if (change < 1e-11)
                break;
        }

        var q = 0.0;
        for (var c = 0; c <= cw; c++)
            q += (c == 0 || c == cw ? 0.5 : 1) * (h1 - h[1, c]);
        return q;
    }

    /// <summary>q/(k·ΔH) for a wall to the depth d in a layer T thick, of any width (Polubarinova-Kochina; Harr).</summary>
    private static double ExactShapeFactor(double t, double d)
    {
        static double K(double m)
        {
            double a = 1, b = Math.Sqrt(1 - m * m);
            while (Math.Abs(a - b) > 1e-15)
                (a, b) = ((a + b) / 2, Math.Sqrt(a * b));
            return Math.PI / (2 * a);
        }
        var alpha = Math.PI * d / (2 * t);
        return K(Math.Cos(alpha)) / (2 * K(Math.Sin(alpha)));
    }
}
