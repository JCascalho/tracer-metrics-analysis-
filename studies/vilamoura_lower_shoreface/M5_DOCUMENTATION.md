# M5 — Tracer Transport and Dispersion Metrics

**Vilamoura lower-shoreface fluorescent sand tracer study (Portugal)**

Study-specific adaptation of the previously validated `M5_Vilamoura_FINAL.py` workflow (scientific revision: October 2026). The scientific calculations, weight definitions, coordinate-frame convention and six campaign intervals are retained. This case-study edition adds command-line input/output selection and revised documentation; **it does not change the principal mathematical algorithms**.

## Scope and principal outputs

- Calibrated concentration: `TMcal = Ac / m`, where `Ac` is fluorescent particle image-area fraction and `m` is the laboratory calibration slope.
- Tracer mass recovery by Thiessen polygon, using penetration depth, sediment density and porosity.
- Concentration-weighted tracer centre of mass (TCOM), displacement distance and azimuth.
- Mean centre-of-mass transport velocity and sediment transport rate `Q = Vm × delta_mix`.
- Coordinate-frame projections with the study's 22° setting, consistently applied to velocity, flux and spatial offsets.
- Previous-TCOM-referenced tracer spreading statistics (manuscript Table 2) **distinct from** variance-based diffusion coefficients.
- Central variance-based diffusion estimates following Miller and Komar (1979) (manuscript Table 3).
- Supplementary fixed-injection-point and sampling-grid-area sensitivity diagnostics.

### Important mathematical distinctions

**Previous-TCOM spreading (Table 2):** terms `Sx`, `Sy`, `Sas`, `Scs` are derived using squared `TMcal` weights and sample offsets from the *previous* campaign TCOM. They include centre-of-mass motion and must **not** be interpreted as central-variance diffusion coefficients.

**Central-moment diffusion (Table 3):** `Dx`, `Dy`, `Das`, `Dcs` arise from the change in `TMcal`-weighted spatial *central* variances about each campaign's *own* TCOM, divided by twice the interval duration. `D_total = Dx + Dy = Das + Dcs` before rounding; quadratic combinations are separate diagnostic quantities. Negative interval estimates indicate contraction of the **observed** variance, not negative physical diffusivity.

The C0 injection is approximated by a zero-variance point source only for the initial variance-based interval C0–C2. The physical release occupied a finite area (~7 m²).

**Grid sensitivity:** the comparison using `TMcal × Thiessen polygon area` is **diagnostic only**. It does not replace the manuscript's primary `TMcal`-weighted estimates. Its polygon-constant concentration assumption and sensitivity to variable spatial support should be considered when interpreting the plot.

## Files

- `M5_tracer_transport_dispersion_metrics.py` — executable workflow.
- `M5_DOCUMENTATION.md` — full methodological and availability documentation (this file).
- `examples/INPUT_TRACER_SYNTHETIC.xlsx` — entirely synthetic workbook with the required `CM0`, `C2`–`C7` sheets; no unpublished positions or observations.

## Required input

The original script reads the `CM0` sheet to obtain the injection location, reference date and global constants, and processes sampling sheets `C2`–`C7` in order. Each campaign contains dates, projected coordinates `x`, `y`, image-derived `Ac`, calibration slope `m`, and Thiessen area `Tp_area`. Optional per-campaign values override global constants where applicable. Transport estimates additionally require a populated `delta_mix` (or recognised spelling), imported from the hydrodynamic analysis rather than calculated by M5.

Key fields:

| Column | Interpretation | Typical units |
| --- | --- | --- |
| `Start_date`, `End_date` | Campaign / injection interval dates | datetime |
| `x`, `y` | Projected sampling coordinates | m |
| `Ac` | Fluorescent tracer image-area fraction | dimensionless |
| `m` | Laboratory calibration slope | calibration-specific |
| `Tp_area` | Thiessen polygon represented area | m² |
| `tracer_nr` | Detected tracer grain count | count |
| `photo_area` | Photographed sampling area | m² |
| `delta_mix` | Active mixing-layer thickness, supplied externally | m |
| `dreger_d` | Sediment sampling penetration depth | m |
| `rho_sed` | Sediment density | kg m⁻³ |
| `sed_poro` | Sediment porosity | dimensionless |
| `tracer_m` | Injected tracer mass | kg |

`CM0` contains `x`, `y`, `End_date` and global constants. The original script uses the first row in each campaign to determine the representative end date and campaign constants. All intervals should have positive elapsed time and positive summed `TMcal`.

## Synthetic demonstration

Install `pandas`, `openpyxl` and `matplotlib`; run from a working directory where the outputs should be written:

```bash
python -m pip install pandas openpyxl matplotlib
python M5_tracer_transport_dispersion_metrics.py --input examples/INPUT_TRACER_SYNTHETIC.xlsx --output M5_SYNTHETIC_RESULTS.xlsx
```

The script generates `Tracer_Recovery_Histogram.png` and `Sampling_Grid_Sensitivity.png` in the working directory, replacing files with the same names. The example is **artificial** and must never be interpreted as Vilamoura experimental results. The graphs provided by the study authors are **not** included as example outputs because they are generated from unpublished observational data.

For private reproduction of the research calculations:

```bash
python M5_tracer_transport_dispersion_metrics.py --input input_TRACER_DATA_delta_mix_p95.xlsx --output M5_TRACER_TRANSPORT_DISPERSION_RESULTS.xlsx
```

## Output workbook

- `CM0`, `C2`–`C7`, `CM2`–`CM7` — reference, calculated sample data and TCOM campaign summaries.
- `TCOM_Displacements` — TCOM intervals, displacements and azimuths.
- `Vm_Q_table_10^(-8)` — velocities and transport rates (numeric values displayed after scaling by 10⁸).
- `Tracer_Spreading_10^-8` — previous-TCOM spreading measures, manuscript Table 2.
- `Variance_Based_Diffusion` — central variance change, manuscript Table 3.
- `Tracer_Recovery_Summary` — recovered tracer mass and percentage by campaign.
- `IP_Referenced_Dispersion`, `Sampling_Grid_Sensitivity`, `Grid_Diffusion_Comparison` — supplementary diagnostics.

The manuscript's Table 2 and Table 3 coefficient columns are expressed in **10⁻⁸ m² s⁻¹**. Labels `as` / `cs` refer to the study-aligned alongshore and cross-shelf components. The quadratic diagnostic resultants in CM sheets must not be substituted for `D_total`.

## Data availability

The complete Vilamoura observational input workbook, numerical results and experimental figures are **not deposited in this public code release**. They are intended to be deposited following publication of the associated research article, subject to journal requirements. The supplied workbook is synthetic.

## Scientific provenance

- Original method and scientific implementation: João Cascalho, FCUL / Instituto Dom Luiz, University of Lisbon, Portugal.
- The central-moment diffusion analysis follows Miller and Komar (1979); see the article bibliography for the full source.
- The 95th-percentile-based `delta_mix` used in the original input is calculated by the preceding M4 near-bed hydrodynamic study and is **read**, not recalculated, by M5.
- Repository: https://github.com/JCascalho/tracer-metrics-analysis
- Cite the immutable tagged release and Zenodo archive after publication of the case-study release; do not cite `main` as an immutable snapshot.
