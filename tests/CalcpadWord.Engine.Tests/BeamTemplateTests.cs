namespace CalcpadWord.Engine.Tests;

/// <summary>
/// The Dokimos RC beam template's script (templates/dokimos/Template source/beam, linked as
/// Fixtures/RcBeamEc2.cpd), calculated as the add-in does in Word.
/// </summary>
public class BeamTemplateTests
{
    private static readonly TemplateScript Beam = new("RcBeamEc2.cpd");

    private static Calculation Calculate(Dictionary<string, string>? values = null, params string[] probes) =>
        Beam.Calculate(values, probes: probes);

    [Fact]
    public void Has_input_fields_for_everything_a_designer_sets() =>
        Assert.Equal(
            ["f_ck", "α_cc", "d_g", "f_yk", "c_min,dur", "RH", "R", "L", "P_b", "b", "h", "c_nom", "n_t", "Φ_t", "n_c", "Φ_c", "n_w", "Φ_w", "s", "g_k,s", "q_k", "ψ_2"],
            Beam.InputNames);

    [Fact]
    public void The_default_beam_calculates_without_errors_and_passes()
    {
        var c = Calculate();

        Assert.False(c.HasErrors);
        Assert.Contains("The beam passes every check.", c.LastBlock);
        Assert.DoesNotContain("FAIL", c.LastBlock);
    }

    [Fact]
    public void Values_match_a_hand_calculation()
    {
        var c = Calculate();

        // w_Ed = 1.35·(10 + 25·0.3·0.5) + 1.5·8 = 30.5625 kN/m over 6 m
        Assert.Equal(30.56, c.Value("w_Ed"), 2);
        Assert.Equal(137.53, c.Value("M_Ed"), 2);
        Assert.Equal(91.69, c.Value("V_Ed"), 2);
        // d = 500 − 30 − 8 − 20/2
        Assert.Equal(452, c.Value("d"), 2);
        Assert.Equal(17, c.Value("f_cd"), 2);
        // V_Rd,c = 0.12·k·(100·ρ_l·f_ck)^(1/3)·b·d, k = 1 + √(200/452), ρ_l = 942.5/(300·452)
        Assert.Equal(74.57, c.Value("V_Rd,c"), 1);
        Assert.Equal(13.27, c.Value("l_d"), 2);
        // Cover, XC1: max(15; 20 − 8) + 10 = 25 mm
        Assert.Equal(25, c.Value("c_nom,req"), 2);
        // Creep (Annex B), RH 50 %, h_0 = 187.5 mm, f_cm = 38: φ_RH = 1.79, β(f_cm) = 2.73, β(28) = 0.488
        Assert.Equal(2.39, c.Value("φ"), 2);
        // Uncracked: I_I = 4.02·10⁹ mm⁴, y = 266.3 mm, so M_cr = 2.9·I_I/(500 − 266.3) = 49.9 kNm
        Assert.Equal(50, c.Value("M_cr"), 0);
        // ζ = 1 − 0.5·(49.9/72.7)² = 0.765; κ = 0.765·4.61·10⁻⁶ + 0.235·2.16·10⁻⁶; δ = 5/48·κ·6000² ≈ 15 mm
        Assert.Equal(15, c.Value("δ_qp"), 0);
    }

    /// <summary>
    /// Tekla Tedds 28.1, "RC beam design (EN1992)", run on 2026-11-18 for the default section
    /// (M = 137.5 kNm, V = 91.7 kN, 3 H20 bottom, 2 H12 top, H8 links at 200) with the UK
    /// annex, whose α_cc = 0.85 is this template's default. Tedds rounds to whole millimetres.
    /// </summary>
    [Fact]
    public void Agrees_with_Tedds_at_its_default_long_term_coefficient()
    {
        var c = Calculate(null, "z", "x", "A_s1,req");

        Assert.Equal(0.075, c.Value("K"), 3);
        Assert.Equal(420, c.Probed("z"), 0);
        Assert.Equal(80, c.Probed("x"), 0);
        Assert.Equal(753, c.Probed("A_s1,req"), 0);
        Assert.Equal(204, c.Value("A_s,min"), 0);
        Assert.Equal(6000, c.Value("A_s,max"), 0);
        Assert.Equal(82, c.Value("s_b"), 0);
        Assert.Equal(200, c.Value("s_c"), 0);
        Assert.Equal(25, c.Value("s_b,min"), 0);
        Assert.Equal(0.263, c.Value("A_sw,min"), 3); // Tedds: 263 mm²/m
        Assert.Equal(0.503, c.Value("A_sw"), 3); // Tedds: 503 mm²/m
        Assert.Equal(339, c.Value("s_l,max"), 0);
    }

    /// <summary>The same Tedds calculation with the recommended values (α_cc = 1.0, k_1 = 0.44, k_2 = 1.25).</summary>
    [Fact]
    public void Agrees_with_Tedds_with_the_recommended_values()
    {
        var c = Calculate(new() { ["α_cc"] = "1" }, "z", "x", "A_s1,req");

        Assert.Equal(0.196, c.Value("K′"), 3);
        Assert.Equal(425, c.Probed("z"), 0);
        Assert.Equal(67, c.Probed("x"), 0);
        Assert.Equal(744, c.Probed("A_s1,req"), 0);
    }

    [Fact]
    public void A_shallow_beam_fails_and_says_where()
    {
        var c = Calculate(new() { ["h"] = "300", ["L"] = "8" });

        Assert.Contains("FAIL", c.LastBlock);
        Assert.Contains("The beam fails at least one check", c.LastBlock);
        // A design that fails is not a calculation that went wrong: no Calcpad errors.
        Assert.False(c.HasErrors);
    }

    [Fact]
    public void Severe_exposure_and_a_narrow_beam_fail_cover_and_fire_resistance()
    {
        // XD3/XS3 (45 mm for durability) with 30 mm cover; 150 mm wide at R120 (Table 5.5 needs 200 mm).
        var c = Calculate(new() { ["c_min,dur"] = "45", ["R"] = "120", ["b"] = "150", ["n_t"] = "2" });

        Assert.Equal(55, c.Value("c_nom,req"), 0);
        Assert.Equal(0.3, c.Value("w_max"), 2);
        Assert.True(c.Value("η_cov") > 1);
        Assert.True(c.Value("η_fi") > 1);
        Assert.Contains("The beam fails at least one check", c.LastBlock);
        Assert.False(c.HasErrors);
    }

    [Fact]
    public void Changing_an_input_changes_the_results() =>
        Assert.True(Calculate(new() { ["q_k"] = "16" }).Value("M_Ed") > Calculate().Value("M_Ed"));
}
