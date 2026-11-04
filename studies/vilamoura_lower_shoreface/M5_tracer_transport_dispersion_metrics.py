"""
====================================================================================
M5 — TRACER TRANSPORT AND DISPERSION METRICS | VILAMOURA LOWER SHOREFACE
File: M5_tracer_transport_dispersion_metrics.py
====================================================================================

Author: João Cascalho (FCUL / Instituto Dom Luiz, University of Lisbon)
Original development: December 2025
Scientific formulation revised: October 2026
Case-study packaging: November 2026
Manuscript: Continental Shelf Research, CSR-D-26-00173

Purpose
-------
Process fluorescent sand-tracer observations from the Vilamoura lower shoreface
and calculate tracer recovery, displacement, transport and spatial spreading.
The two principal spatial metrics correspond to separate manuscript tables:

    Table 2 — Tracer spreading metrics (Sx, Sy, Sas, Scs)
        Squared calibrated-concentration weighting, using sample offsets from
        the PREVIOUS campaign's tracer centre of mass (TCOM). These metrics
        include the effect of TCOM displacement and are NOT variance-based
        diffusion coefficients. Combined quadratic magnitudes are reported.
        Excel sheet: Tracer_Spreading_10^-8.

    Table 3 — Variance-based diffusion estimates (Dx, Dy, Das, Dcs, D_total)
        Method of moments following Miller and Komar (1979): changes in
        TMcal-weighted CENTRAL spatial variance, calculated about each
        campaign's own TCOM, divided by twice the elapsed interval time.
        D_total = Dx + Dy = Das + Dcs (before rounding). Signed negative
        values describe reductions in OBSERVED tracer variance, not negative
        physical diffusivity.
        Excel sheet: Variance_Based_Diffusion.

The C0 release is approximated as a point source with zero initial variance
ONLY for the first variance-based diffusion interval (C0–C2). The physical
release occupied a finite seabed area (approximately 7 m²).

Additional calculations
-----------------------
    • Image-derived calibrated tracer concentration (TMcal = Ac / m)
    • TCOM coordinates, successive displacements and transport velocities
    • Alongshore/cross-shelf projections using a 22° reference-frame rotation
    • Horizontal tracer transport fluxes and active mixing-layer thickness
      read from the input data
    • Tracer grain counts per square metre
    • Tracer mass recovery using ArcGIS Pro Thiessen polygon areas
    • Fixed injection-point second moments and cumulative/interval dispersion
      (supplementary diagnostics; include centre-of-mass displacement)
    • Sampling-grid sensitivity based on TMcal × Thiessen-area weights
      (supplementary diagnostics; does not replace the principal estimates)

Scientific basis and recovery calculation
-----------------------------------------
Tracer pixels in a fixed photographic area provide the area fraction:

    Ac = tracer-particle pixels / total photograph pixels

Laboratory calibration gives Ac = m × TMc; thus TMcal = Ac / m, where m is
the calibration slope. The calibrated concentration TMcal is used as the
weight for the principal tracer-centre and spatial-moment calculations.

Tracer mass inventoried in each Thiessen polygon is estimated as:

    T_rate = TMcal × Tp_area × dreger_d × rho_sed × (1 - sed_poro)
    T_rate_% = 100 × T_rate / tracer_m

where Tp_area is polygon area (m²), dreger_d is sediment penetration depth
(m), rho_sed is sediment density (kg/m³), sed_poro is sediment porosity,
and tracer_m is the injected tracer mass (approximately 600 kg).

Input and output
----------------
Input: input_TRACER_DATA_delta_mix_p95.xlsx
    • CM0: injection reference coordinates, dates and shared constants
    • C2–C7: sampling observations, positions, calibration and polygon areas
    • Start_date / End_date: dates used for elapsed-time calculations
    • Positive total TMcal required for each campaign's spatial moments

Output: Vilamoura_M5_Results_FINAL.xlsx
    • C2–C7 and CM2–CM7: detailed campaign calculations and summaries
    • Tracer_Spreading_10^-8: manuscript Table 2
    • Variance_Based_Diffusion: manuscript Table 3, including D_total
    • TCOM_Displacements, Vm_Q_table_10^(-8), Tracer_Recovery_Summary
    • IP_Referenced_Dispersion, Sampling_Grid_Sensitivity and
      Grid_Diffusion_Comparison: supplementary diagnostic analyses
    • Tracer_Recovery_Histogram.png and Sampling_Grid_Sensitivity.png

Table 2 and Table 3 spatial metrics are scaled by 10^-8 m² s^-1 in the
output worksheets. The labels 'as' and 'cs' mean alongshore and cross-shelf.
Quadratic resultants retained in individual CM sheets are diagnostics and
must not be substituted for D_total.

Reproducibility
---------------
Run with Python 3, pandas and openpyxl; matplotlib is used for PNG plots.
Keep the input workbook alongside this script. Running the script overwrites
output files with the same names in the current working directory.
See M5_METHODS.md for the manuscript-table correspondence and data policy.

====================================================================================
"""
import pandas as pd
import math
import argparse
from openpyxl import load_workbook


# =============================================================================
# USER CONFIGURATION
# =============================================================================
# Angle used to project the Cartesian COM velocity vector (Vm_x, Vm_y)
# into the alongshore/cross-shelf reference frame.
#
# IMPORTANT:
# This is NOT a geometric rotation of the vector, but a change of reference frame.
# The sign of the angle defines the orientation of the positive alongshore axis.
#
# The formula below projects Cartesian components into a shoreline-oriented
# frame. With VELOCITY_ROTATION_DEG = +22°, positive alongshore transport
# is defined approximately toward ESE/SE for the Belharucas/Falésia coast.
#
# If the alongshore sign convention is reversed, use +202° (= 22° + 180°)
# or equivalently −158°.
#
# The same angle is also used to project sample offsets (Dist_x, Dist_y)
# into Dist_as and Dist_cs for diffusion calculations. Therefore, diffusion
# in the alongshore/cross-shelf frame no longer depends on pre-existing
# as/cs columns in the input workbook.
VELOCITY_ROTATION_DEG = 22.0


def rotate_vector_xy_to_ascs(vx: float, vy: float, angle_deg: float = VELOCITY_ROTATION_DEG):
    """
    Project a Cartesian vector (vx, vy) into alongshore/cross-shelf components.

    This is used for transport velocities and fluxes. It avoids deriving
    Vm_as/Vm_cs from differences in pre-rotated point coordinates, which can
    make the transport sign sensitive to how as/cs were generated upstream.

    The sign convention is controlled by VELOCITY_ROTATION_DEG.
    """
    if pd.isna(vx) or pd.isna(vy):
        return float("nan"), float("nan")

    theta = math.radians(angle_deg)
    v_as = vx * math.cos(theta) - vy * math.sin(theta)
    v_cs = vx * math.sin(theta) + vy * math.cos(theta)
    return v_as, v_cs


def project_xy_to_ascs(dx, dy, angle_deg: float = VELOCITY_ROTATION_DEG):
    """
    Project Cartesian offsets (dx, dy) into alongshore/cross-shelf offsets.

    Used for diffusion calculations so that D_as and D_cs are computed from
    Dist_x and Dist_y using the same reference-frame convention as Vm_as/Vm_cs.
    This avoids dependence on pre-existing as/cs coordinate columns.
    """
    theta = math.radians(angle_deg)
    d_as = dx * math.cos(theta) - dy * math.sin(theta)
    d_cs = dx * math.sin(theta) + dy * math.cos(theta)
    return d_as, d_cs


def scaled_1e8(val: float) -> float:
    """
    Scale value by 1e8 and round to 2 decimals.
    Keep NaN as NaN (Excel will show empty).
    """
    if pd.isna(val):
        return float("nan")
    return round(val * 1e8, 2)


def main():
    # -------------------------------------------------
    # 0. File names (change if needed)
    # -------------------------------------------------
    parser = argparse.ArgumentParser(
        description="M5 — Vilamoura tracer recovery, transport and dispersion metrics"
    )
    parser.add_argument(
        "--input", default="input_TRACER_DATA_delta_mix_p95.xlsx",
        help="Excel workbook with CM0 and C2–C7 worksheets",
    )
    parser.add_argument(
        "--output", default="M5_TRACER_TRANSPORT_DISPERSION_RESULTS.xlsx",
        help="Output Excel workbook (existing file will be overwritten)",
    )
    args = parser.parse_args()
    input_file = args.input
    output_file = args.output

    # -------------------------------------------------
    # 1. Load workbook and read CM0 (reference sheet)
    # -------------------------------------------------
    xls = pd.ExcelFile(input_file)

    if "CM0" not in xls.sheet_names:
        raise ValueError("Sheet 'CM0' not found! It is required as reference.")

    cM0 = pd.read_excel(
        xls, sheet_name="CM0", parse_dates=["Start_date", "End_date"]
    )

    # Reference end date and CM0 coordinates (for first Δt and first displacement)
    previous_end_date = cM0.loc[0, "End_date"]
    prev_CMx = cM0.loc[0, "x"]
    prev_CMy = cM0.loc[0, "y"]

    # Use CM0 as the origin for diagnostic alongshore/cross-shelf COM positions.
    # These are computed internally from x/y and are not read from input as/cs columns.
    as_origin_x = prev_CMx
    as_origin_y = prev_CMy
    prev_CMas = 0.0
    prev_CMcs = 0.0

    # ---- global constants (read from CM0 if present, else defaults) ----
    # drag penetration depth [m]
    if "dreger_d" in cM0.columns:
        dreger_d_global = float(cM0["dreger_d"].iloc[0])
    else:
        dreger_d_global = 0.15

    # sediment porosity [-]
    if "sed_poro" in cM0.columns:
        sed_poro_global = float(cM0["sed_poro"].iloc[0])
    else:
        sed_poro_global = 0.6

    # sediment density [kg/m3] – will be overridden per sheet if present
    if "rho_sed" in cM0.columns:
        rho_sed_global = float(cM0["rho_sed"].iloc[0])
    else:
        rho_sed_global = float("nan")

    # injected tracer mass [kg]
    if "tracer_m" in cM0.columns:
        tracer_m_total = float(cM0["tracer_m"].iloc[0])
    else:
        tracer_m_total = 600.0

    # -------------------------------------------------
    # 2. Define the sheets to process (C2 to C7)
    # -------------------------------------------------
    target_sheets = [f"C{i}" for i in range(2, 8)]  # C2, C3, C4, C5, C6, C7

    # Containers to store results
    processed_sheets = {}   # updated C2..C7
    com_sheets = {}         # CM2..CM7
    vmq_rows = []           # rows for Vm/Q summary table (scaled numeric)
    diff_rows = []          # rows for diffusion summary table (scaled numeric)
    tcom_rows = []          # rows for successive TCOM displacement summary table
    prev_label = "C0"       # for interval labels C0–C2, C2–C3, ...
    # C0 is the injection point (IP), idealized as a point source.
    # Hence initial central variances are zero along every axis.
    # This assumption does not describe the finite ~7 m² physical release patch.
    prev_variances = {k: 0.0 for k in ("x", "y", "as", "cs")}
    spreading_rows = []
    ip_dispersion_rows = []
    grid_sensitivity_rows = []
    grid_interval_rows = []
    prev_grid_variances = {k: 0.0 for k in ("x", "y", "as", "cs")}
    prev_ip_moments = {k: 0.0 for k in ("x", "y", "as", "cs")}
    injection_date = previous_end_date

    # -------------------------------------------------
    # 3. Loop over C2–C7 and compute metrics
    # -------------------------------------------------
    for sheet_name in target_sheets:

        if sheet_name not in xls.sheet_names:
            print(f"⚠ WARNING: Sheet '{sheet_name}' not found in workbook. Skipping.")
            continue

        print(f"\n📌 Processing sheet: {sheet_name}")

        df = pd.read_excel(
            xls, sheet_name=sheet_name, parse_dates=["Start_date", "End_date"]
        )

        if "End_date" not in df.columns:
            raise KeyError(f"'End_date' column not found in sheet '{sheet_name}'.")

        # End_date for this campaign (assume representative = first row)
        this_end_date = df["End_date"].iloc[0]

        # ---------------------------------------------
        # 3.1 Compute Δt [seconds] between SUCCESSIVE CM's
        # ---------------------------------------------
        delta_t_seconds = (this_end_date - previous_end_date).total_seconds()
        df["delta_t"] = delta_t_seconds

        # ---------------------------------------------
        # 3.1a Tracer grains per square metre
        #       grains_per_m2 = tracer_nr / photo_area
        # ---------------------------------------------
        if "tracer_nr" in df.columns and "photo_area" in df.columns:
            df["tracer_grains_m2"] = df["tracer_nr"] / df["photo_area"]
        else:
            print(
                f"⚠ WARNING: 'tracer_nr' or 'photo_area' not found in '{sheet_name}'. "
                "Column 'tracer_grains_m2' will not be created."
            )
            df["tracer_grains_m2"] = float("nan")

        # ---------------------------------------------
        # 3.2 Compute TMcal = Ac / m
        # ---------------------------------------------
        if "Ac " in df.columns:
            ac_col = "Ac "
        elif "Ac" in df.columns:
            ac_col = "Ac"
        else:
            raise KeyError(f"Missing 'Ac' or 'Ac ' column in sheet '{sheet_name}'.")

        df["TMcal"] = df[ac_col] / df["m"]

        # Scaled nondimensional tracer concentration for ArcGIS mapping.
        # Keeps TMcal unchanged for all scientific calculations.
        df["TMcal_10000"] = df["TMcal"] * 10000.0

        # ---------------------------------------------
        # 3.2a Tracer mass recuperation per polygon (T_rate, kg)
        #       T_rate = TMcal * Tp_area * dreger_d * rho_sed * (1 - sed_poro)
        #
        # IMPORTANT:
        # sed_poro is porosity, so the solid sediment fraction is (1 - sed_poro).
        # Using sed_poro directly overestimates the recovered mass.
        # ---------------------------------------------
        # Thiessen polygon area
        if "Tp_area" in df.columns:
            tp_area_col = "Tp_area"
        else:
            raise KeyError(
                f"Missing 'Tp_area' column in sheet '{sheet_name}' "
                "(Thiessen polygon area required for T_rate)."
            )

        # get campaign-specific constants if present, else global
        if "dreger_d" in df.columns:
            dreger_d = float(df["dreger_d"].iloc[0])
        else:
            dreger_d = dreger_d_global

        if "sed_poro" in df.columns:
            sed_poro = float(df["sed_poro"].iloc[0])
        else:
            sed_poro = sed_poro_global

        if "rho_sed" in df.columns:
            rho_sed = float(df["rho_sed"].iloc[0])
        else:
            if math.isnan(rho_sed_global):
                raise KeyError(
                    f"'rho_sed' (sediment density) not found in sheet '{sheet_name}' "
                    "or in CM0. Please add it to the input file."
                )
            rho_sed = rho_sed_global

        # tracer mass injected (same for all campaigns)
        tracer_m = tracer_m_total

        solid_fraction = 1.0 - sed_poro
        if solid_fraction < 0 or solid_fraction > 1:
            raise ValueError(
                f"Invalid sed_poro value in sheet {sheet_name}: {sed_poro}. "
                "Porosity must be between 0 and 1."
            )

        df["T_rate"] = (
            df["TMcal"] * df[tp_area_col] * dreger_d * rho_sed * solid_fraction
        )
        df["T_rate_%"] = 100.0 * df["T_rate"] / tracer_m

        # ---------------------------------------------
        # 3.3 Distances from previous center of mass (per sample)
        #      Dist_x(i) = x_i - CMx_prev, etc.
        # ---------------------------------------------
        df["Dist_x"] = df["x"] - prev_CMx
        df["Dist_y"] = df["y"] - prev_CMy

        # Project sample offsets into the alongshore/cross-shelf frame using
        # the same angle used for velocity/transport. This avoids dependence
        # on pre-rotated as/cs columns in the input workbook.
        df[["Dist_as", "Dist_cs"]] = df.apply(
            lambda r: project_xy_to_ascs(
                r["Dist_x"], r["Dist_y"], VELOCITY_ROTATION_DEG
            ),
            axis=1,
            result_type="expand",
        )

        # For checking like your Excel example:
        df["TMcal*Dist_x"] = df["TMcal"] * df["Dist_x"]
        df["TMcal*Dist_y"] = df["TMcal"] * df["Dist_y"]
        df["TMcal*Dist_as"] = df["TMcal"] * df["Dist_as"]
        df["TMcal*Dist_cs"] = df["TMcal"] * df["Dist_cs"]

        df["(TMcal*Dist_x)^2"] = df["TMcal*Dist_x"] ** 2
        df["(TMcal*Dist_y)^2"] = df["TMcal*Dist_y"] ** 2
        df["(TMcal*Dist_as)^2"] = df["TMcal*Dist_as"] ** 2
        df["(TMcal*Dist_cs)^2"] = df["TMcal*Dist_cs"] ** 2

        # ---------------------------------------------
        # 3.4 TMcal-weighted coordinates for COM
        # ---------------------------------------------
        df["TMcal*x"] = df["TMcal"] * df["x"]
        df["TMcal*y"] = df["TMcal"] * df["y"]

        # ---------------------------------------------
        # 3.5 Compute Center of Mass for this sheet
        # ---------------------------------------------
        TM_sum = df["TMcal"].sum()
        TM_sq_sum = (df["TMcal"] ** 2).sum()

        # Independent grid-sensitivity analysis (diagnostic only):
        # w_i = TMcal_i * Thiessen_area_i. No changes to the original
        # TMcal-weighted COM, diffusion, velocities, fluxes or recovery.
        grid_cm = {k: float("nan") for k in ("x", "y", "as", "cs")}
        grid_variances = {k: float("nan") for k in ("x", "y", "as", "cs")}
        grid_D = {k: float("nan") for k in ("x", "y", "as", "cs")}
        grid_weight_sum = float("nan")
        valid_grid = False

        # Initialise all metrics
        D_x = D_y = D_as = D_cs = float("nan")
        spreading_S = {k: float("nan") for k in ("x", "y", "as", "cs")}
        variances = {k: float("nan") for k in ("x", "y", "as", "cs")}
        D_xy_combined = D_as_cs_combined = float("nan")
        ip_moments = {k: float("nan") for k in ("x", "y", "as", "cs")}
        ip_cumulative = {k: float("nan") for k in ("x", "y", "as", "cs")}
        ip_interval = {k: float("nan") for k in ("x", "y", "as", "cs")}
        ip_elapsed_seconds = (this_end_date - injection_date).total_seconds()
        Vm_x = Vm_y = Vm_as = Vm_cs = float("nan")
        Q_x = Q_y = Q_as = Q_cs = float("nan")
        delta_mix_val = float("nan")
        grains_m2_mean = float("nan")
        T_rate_sum = float("nan")
        T_rate_pct_sum = float("nan")

        if TM_sum == 0 or math.isclose(TM_sum, 0.0):
            print(
                f"⚠ WARNING: Sum(TMcal) = 0 in '{sheet_name}'. "
                "COM, diffusion & transport set to NaN."
            )
            CMx = CMy = CMas = CMcs = float("nan")
        else:
            # ----- Center of mass -----
            CMx = df["TMcal*x"].sum() / TM_sum
            CMy = df["TMcal*y"].sum() / TM_sum

            # Diagnostic COM position in the alongshore/cross-shelf frame,
            # computed internally from Cartesian COM coordinates relative to CM0.
            CMas, CMcs = project_xy_to_ascs(
                CMx - as_origin_x,
                CMy - as_origin_y,
                VELOCITY_ROTATION_DEG,
            )

            # Mean grains/m2 in this campaign (simple arithmetic mean)
            grains_m2_mean = df["tracer_grains_m2"].mean()

            # Total tracer mass recovered/inventoried in this campaign
            T_rate_sum = df["T_rate"].sum()
            T_rate_pct_sum = 100.0 * T_rate_sum / tracer_m

            if T_rate_pct_sum > 100.0:
                print(
                    f"⚠ WARNING: {sheet_name} recovery is {T_rate_pct_sum:.2f}% "
                    "of injected tracer mass. Check TMcal calibration, Tp_area, "
                    "dreger_d, rho_sed, and sed_poro."
                )

            # -----------------------------------------
            # 3.6 Method-of-moments spreading (Miller & Komar, 1979)
            # Variances are computed about the CURRENT campaign COM,
            # with linear TMcal weights. Diffusion = variance increment/(2*dt).
            # C0 is idealized as a point source, so its initial variance is zero.
            # Therefore the C0-C2 diffusion estimate is also calculated.
            # -----------------------------------------
            current_offsets = {
                "x": df["x"] - CMx,
                "y": df["y"] - CMy,
            }
            current_offsets["as"], current_offsets["cs"] = project_xy_to_ascs(
                current_offsets["x"], current_offsets["y"], VELOCITY_ROTATION_DEG
            )
            for axis, offsets in current_offsets.items():
                variances[axis] = float((df["TMcal"] * offsets**2).sum() / TM_sum)
                df[f"Offset_current_CM_{axis}"] = offsets

            if prev_variances is not None and delta_t_seconds > 0:
                D_x, D_y, D_as, D_cs = (
                    (variances[k] - prev_variances[k]) / (2 * delta_t_seconds)
                    for k in ("x", "y", "as", "cs")
                )
                # Resultant of signed diffusion components is a diagnostic,
                # not a tensor-invariant diffusion magnitude.
                D_xy_combined = math.hypot(D_x, D_y)
                D_as_cs_combined = math.hypot(D_as, D_cs)

            # -----------------------------------------
            # 3.6a SOURCE-REFERENCED second moments about the fixed IP (CM0).
            # These include BOTH central spreading and squared COM displacement.
            # M2_IP = Var_current_CM + (CM_current - IP)^2.
            # Use linear TMcal weights, consistently with central moments.
            # K_IP_cumulative = M2_IP / (2 * elapsed time since injection).
            # K_IP_interval = (M2_IP_now - M2_IP_previous) / (2 * interval dt).
            # Negative interval values remain signed; no absolute value applied.
            # Neither metric is an advection-independent diffusion coefficient.
            # -----------------------------------------
            ip_offsets = {
                "x": df["x"] - as_origin_x,
                "y": df["y"] - as_origin_y,
            }
            ip_offsets["as"], ip_offsets["cs"] = project_xy_to_ascs(
                ip_offsets["x"], ip_offsets["y"], VELOCITY_ROTATION_DEG
            )
            for axis, offsets in ip_offsets.items():
                ip_moments[axis] = float((df["TMcal"] * offsets**2).sum() / TM_sum)
                df[f"Offset_IP_{axis}"] = offsets
                if ip_elapsed_seconds > 0:
                    ip_cumulative[axis] = ip_moments[axis] / (2 * ip_elapsed_seconds)
                if delta_t_seconds > 0:
                    ip_interval[axis] = (
                        ip_moments[axis] - prev_ip_moments[axis]
                    ) / (2 * delta_t_seconds)

            # Check the exact second-moment decomposition numerically.
            com_displacements = {
                "x": CMx - as_origin_x,
                "y": CMy - as_origin_y,
                "as": CMas,
                "cs": CMcs,
            }
            for axis in ("x", "y", "as", "cs"):
                expected = variances[axis] + com_displacements[axis]**2
                if not math.isclose(ip_moments[axis], expected, rel_tol=1e-7, abs_tol=1e-5):
                    raise ValueError(f"IP moment identity failed for {sheet_name}/{axis}")

            # -----------------------------------------
            # 3.6b Grid sensitivity: area-weighted spatial moments
            # (additional diagnostics; primary metrics remain unchanged).
            # Each Thiessen polygon represents a spatial support area.
            # The approximation assumes constant TMcal within each polygon.
            # -----------------------------------------
            area = pd.to_numeric(df[tp_area_col], errors="coerce")
            conc = pd.to_numeric(df["TMcal"], errors="coerce")
            coords_ok = df[["x", "y"]].notna().all(axis=1)
            invalid = (~area.notna() | ~conc.notna() | ~coords_ok |
                       (area <= 0) | (conc < 0))
            if invalid.any():
                print(f"⚠ WARNING: {sheet_name}: {int(invalid.sum())} invalid records "
                      "for grid sensitivity; area-weighted moments left undefined.")
            else:
                grid_weights = conc * area
                grid_weight_sum = float(grid_weights.sum())
                if grid_weight_sum > 0:
                    valid_grid = True
                    grid_cm["x"] = float((grid_weights * df["x"]).sum() / grid_weight_sum)
                    grid_cm["y"] = float((grid_weights * df["y"]).sum() / grid_weight_sum)
                    grid_cm["as"], grid_cm["cs"] = project_xy_to_ascs(
                        grid_cm["x"] - as_origin_x,
                        grid_cm["y"] - as_origin_y,
                        VELOCITY_ROTATION_DEG,
                    )
                    grid_offsets = {
                        "x": df["x"] - grid_cm["x"],
                        "y": df["y"] - grid_cm["y"],
                    }
                    grid_offsets["as"], grid_offsets["cs"] = project_xy_to_ascs(
                        grid_offsets["x"], grid_offsets["y"], VELOCITY_ROTATION_DEG
                    )
                    for axis, offsets in grid_offsets.items():
                        grid_variances[axis] = float(
                            (grid_weights * offsets**2).sum() / grid_weight_sum
                        )
                    if prev_grid_variances is not None and delta_t_seconds > 0:
                        grid_D = {
                            k: (grid_variances[k] - prev_grid_variances[k]) /
                               (2 * delta_t_seconds)
                            for k in ("x", "y", "as", "cs")
                        }
                else:
                    print(f"⚠ WARNING: {sheet_name}: zero total area-weighted concentration.")

            # Previous-COM-referenced tracer spreading metric (Table 2).
            # This is not the variance-based method-of-moments diffusion estimate.
            if delta_t_seconds > 0 and TM_sq_sum > 0:
                for axis in ("x", "y", "as", "cs"):
                    spreading_S[axis] = float(
                        df[f"(TMcal*Dist_{axis})^2"].sum()
                        / (TM_sq_sum * delta_t_seconds)
                    )

            # -----------------------------------------
            # 3.7 Mean tracer displacement velocities
            #     based on Cartesian COM displacement.
            #
            # IMPORTANT CHANGE:
            # Vm_as and Vm_cs are no longer computed from
            # (CM_as - prev_CMas) / delta_t and
            # (CM_cs - prev_CMcs) / delta_t.
            #
            # Instead, the physically meaningful Cartesian COM velocity vector
            # (Vm_x, Vm_y) is rotated by VELOCITY_ROTATION_DEG into the
            # alongshore/cross-shelf frame. This makes Q_as and Q_cs vectorial
            # quantities derived consistently from the same displacement vector.
            # -----------------------------------------
            dist_cm_x = CMx - prev_CMx
            dist_cm_y = CMy - prev_CMy

            # Keep these diagnostic distances for output/comparison only.
            # They are NOT used to compute Vm_as/Vm_cs or Q_as/Q_cs.
            dist_cm_as = CMas - prev_CMas
            dist_cm_cs = CMcs - prev_CMcs

            if delta_t_seconds == 0:
                print(
                    f"⚠ WARNING: Δt = 0 in '{sheet_name}'. "
                    "Velocities set to NaN."
                )
            else:
                Vm_x = dist_cm_x / delta_t_seconds
                Vm_y = dist_cm_y / delta_t_seconds

                # Project the Cartesian velocity vector into the alongshore/
                # cross-shelf frame. Change VELOCITY_ROTATION_DEG at the top
                # if the shoreline-axis convention needs adjustment.
                Vm_as, Vm_cs = rotate_vector_xy_to_ascs(
                    Vm_x, Vm_y, VELOCITY_ROTATION_DEG
                )

            # -----------------------------------------
            # 3.8 Transport rates Qx, Qy, Qas, Qcs
            # -----------------------------------------
            if "delta_mix" in df.columns:
                delta_mix_col = "delta_mix"
            elif "δ_mix" in df.columns:
                delta_mix_col = "δ_mix"
            elif "delta_mix " in df.columns:
                delta_mix_col = "delta_mix "
            else:
                delta_mix_col = None

            if delta_mix_col is None:
                print(
                    f"⚠ WARNING: No 'delta_mix' column found in '{sheet_name}'. "
                    "Transport Q set to NaN."
                )
            else:
                delta_mix_series = df[delta_mix_col].dropna()
                if delta_mix_series.empty:
                    print(
                        f"⚠ WARNING: 'delta_mix' column empty in '{sheet_name}'. "
                        "Transport Q set to NaN."
                    )
                else:
                    delta_mix_val = float(delta_mix_series.iloc[0])
                    Q_x = Vm_x * delta_mix_val
                    Q_y = Vm_y * delta_mix_val
                    Q_as = Vm_as * delta_mix_val
                    Q_cs = Vm_cs * delta_mix_val

        # Store updated C* sheet
        processed_sheets[sheet_name] = df

        # ---------------------------------------------
        # 3.9 Build corresponding CM* sheet (CM2, CM3, ...)
        # ---------------------------------------------
        cm_sheet_name = "CM" + sheet_name[1:]     # e.g. C2 -> CM2
        cm_value = cm_sheet_name                  # value in first column (CM2, CM3, ...)

        # COM displacement (for reference)
        dist_cm_x = CMx - prev_CMx
        dist_cm_y = CMy - prev_CMy
        dist_cm_as = CMas - prev_CMas
        dist_cm_cs = CMcs - prev_CMcs

        cm_df = pd.DataFrame({
            "CM":                    [cm_value],          # CM2, CM3, ..., CM7
            "CM_x":                  [CMx],
            "CM_y":                  [CMy],
            "CM_as":                 [CMas],
            "CM_cs":                 [CMcs],
            "TMcal_sum":             [TM_sum],
            "TMcal_sq_sum":          [TM_sq_sum],
            "delta_t_s":             [delta_t_seconds],   # time between this CM and previous CM
            "Velocity_rotation_deg": [VELOCITY_ROTATION_DEG],
            "DistCM_x":              [dist_cm_x],
            "DistCM_y":              [dist_cm_y],
            "DistCM_as":             [dist_cm_as],
            "DistCM_cs":             [dist_cm_cs],
            "TCOMd_xy":              [math.sqrt(dist_cm_x**2 + dist_cm_y**2)],
            "TCOMd_ascs":            [math.sqrt(dist_cm_as**2 + dist_cm_cs**2)],
            "TCOMdir_deg_xy":        [(math.degrees(math.atan2(dist_cm_x, dist_cm_y)) + 360) % 360],
            "TCOMdir_deg_ascs":      [(math.degrees(math.atan2(dist_cm_as, dist_cm_cs)) + 360) % 360],
            "Variance_x_m2":         [variances["x"]],
            "Variance_y_m2":         [variances["y"]],
            "Variance_as_m2":        [variances["as"]],
            "Variance_cs_m2":        [variances["cs"]],
            "IP_M2_x_m2":            [ip_moments["x"]],
            "IP_M2_y_m2":            [ip_moments["y"]],
            "IP_M2_as_m2":           [ip_moments["as"]],
            "IP_M2_cs_m2":           [ip_moments["cs"]],
            "D_x":                   [D_x],
            "D_y":                   [D_y],
            "D_as":                  [D_as],
            "D_cs":                  [D_cs],
            "D_xy_combined":         [D_xy_combined],
            "D_as_cs_combined":      [D_as_cs_combined],
            "Vm_x":                  [Vm_x],
            "Vm_y":                  [Vm_y],
            "Vm_as":                 [Vm_as],
            "Vm_cs":                 [Vm_cs],
            "delta_mix":             [delta_mix_val],
            "Q_x":                   [Q_x],
            "Q_y":                   [Q_y],
            "Q_as":                  [Q_as],
            "Q_cs":                  [Q_cs],
            "tracer_grains_m2_mean": [grains_m2_mean],
            "T_rate_sum":            [T_rate_sum],       # total kg in this campaign
            "T_rate_%_sum":          [T_rate_pct_sum],   # % of injected mass
            "prev_EndDate":          [previous_end_date],
            "End_date":              [this_end_date],
        })

        com_sheets[cm_sheet_name] = cm_df

        # ---------------------------------------------
        # 3.10 Collect Vm/Q row for summary table (scaled ×1e8, numeric)
        #      NOW including resultants of Qxy and Qas_cs
        # ---------------------------------------------
        interval_label = f"{prev_label}–{sheet_name}"  # e.g. C0–C2, C2–C3, ...

        # ---------------------------------------------
        # 3.10a Collect successive TCOM displacement row
        #      Distances between consecutive centers of mass
        #      are kept in real spatial units (m), not scaled.
        # ---------------------------------------------
        if any(pd.isna(v) for v in (dist_cm_x, dist_cm_y)):
            TCOMd_xy = float("nan")
            TCOMdir_deg_xy = float("nan")
        else:
            TCOMd_xy = math.sqrt(dist_cm_x**2 + dist_cm_y**2)
            TCOMdir_deg_xy = (math.degrees(math.atan2(dist_cm_x, dist_cm_y)) + 360) % 360

        if any(pd.isna(v) for v in (dist_cm_as, dist_cm_cs)):
            TCOMd_ascs = float("nan")
            TCOMdir_deg_ascs = float("nan")
        else:
            TCOMd_ascs = math.sqrt(dist_cm_as**2 + dist_cm_cs**2)
            TCOMdir_deg_ascs = (math.degrees(math.atan2(dist_cm_as, dist_cm_cs)) + 360) % 360

        tcom_rows.append({
            "Sampling intervals": interval_label,
            "From": prev_label,
            "To": sheet_name,
            "prev_EndDate": previous_end_date,
            "End_date": this_end_date,
            "delta_t_s": delta_t_seconds,
            "prev_CM_x": prev_CMx,
            "prev_CM_y": prev_CMy,
            "CM_x": CMx,
            "CM_y": CMy,
            "DistCM_x_m": dist_cm_x,
            "DistCM_y_m": dist_cm_y,
            "TCOMd_xy_m": TCOMd_xy,
            "TCOMdir_deg_xy": TCOMdir_deg_xy,
            "prev_CM_as": prev_CMas,
            "prev_CM_cs": prev_CMcs,
            "CM_as": CMas,
            "CM_cs": CMcs,
            "DistCM_as_m": dist_cm_as,
            "DistCM_cs_m": dist_cm_cs,
            "TCOMd_ascs_m": TCOMd_ascs,
            "TCOMdir_deg_ascs": TCOMdir_deg_ascs,
        })

        # Compute unscaled resultants, guarding against NaNs
        if any(math.isnan(v) for v in (Q_x, Q_y)):
            Q_xy_res = float("nan")
        else:
            Q_xy_res = math.sqrt(Q_x**2 + Q_y**2)

        if any(math.isnan(v) for v in (Q_as, Q_cs)):
            Q_as_cs_res = float("nan")
        else:
            Q_as_cs_res = math.sqrt(Q_as**2 + Q_cs**2)

        vmq_rows.append({
            "Sampling intervals":    interval_label,
            "Vmx":                   scaled_1e8(Vm_x),
            "Vmy":                   scaled_1e8(Vm_y),
            "Vmas":                  scaled_1e8(Vm_as),
            "Vmcs":                  scaled_1e8(Vm_cs),
            "Qx":                    scaled_1e8(Q_x),
            "Qy":                    scaled_1e8(Q_y),
            "Qas":                   scaled_1e8(Q_as),
            "Qcs":                   scaled_1e8(Q_cs),
            "Q_xy_resultant":        scaled_1e8(Q_xy_res),
            "Q_as_cs_resultant":     scaled_1e8(Q_as_cs_res),
        })

        # ---------------------------------------------
        # 3.11 Collect variance-based diffusion row (Table 3) (scaled ×1e8, numeric)
        # ---------------------------------------------
        diff_rows.append({
            "Sampling intervals": interval_label,
            "Dx":       scaled_1e8(D_x),
            "Dy":       scaled_1e8(D_y),
            "Das":      scaled_1e8(D_as),
            "Dcs":      scaled_1e8(D_cs),
            "D_total": scaled_1e8(D_x + D_y),
        })

        spreading_rows.append({
            "Sampling intervals": interval_label,
            **{f"S{k}": scaled_1e8(spreading_S[k]) for k in ("x", "y", "as", "cs")},
            "S_xy_combined": scaled_1e8(math.hypot(spreading_S["x"], spreading_S["y"])),
            "S_as_cs_combined": scaled_1e8(math.hypot(spreading_S["as"], spreading_S["cs"])),
        })
        # IP-referenced diagnostics (all coefficients scaled by 10^8).
        # Totals are sums of orthogonal components, not quadratic resultants.
        ip_dispersion_rows.append({
            "Sampling intervals": interval_label,
            "Campaign": sheet_name,
            "Elapsed_since_IP_s": ip_elapsed_seconds,
            "Interval_dt_s": delta_t_seconds,
            **{f"M2_IP_{k}_m2": ip_moments[k] for k in ("x", "y", "as", "cs")},
            "M2_IP_total_m2": ip_moments["x"] + ip_moments["y"],
            **{f"K_IP_cum_{k}_1e-8": scaled_1e8(ip_cumulative[k]) for k in ("x", "y", "as", "cs")},
            "K_IP_cum_total_1e-8": scaled_1e8(ip_cumulative["x"] + ip_cumulative["y"]),
            **{f"K_IP_interval_{k}_1e-8": scaled_1e8(ip_interval[k]) for k in ("x", "y", "as", "cs")},
            "K_IP_interval_total_1e-8": scaled_1e8(ip_interval["x"] + ip_interval["y"]),
            "D_central_total_1e-8": scaled_1e8(D_x + D_y),
        })
        # Campaign-level sensitivity table (unrounded moments in m and m²).
        # The main analysis uses concentrations only; area-weighted estimates
        # are separate and do not overwrite any original output columns.
        grid_sensitivity_rows.append({
            "Campaign": sheet_name,
            "N_samples": len(df),
            "N_positive_TMcal": int((df["TMcal"] > 0).sum()),
            "Thiessen_area_min_m2": float(df[tp_area_col].min()),
            "Thiessen_area_median_m2": float(df[tp_area_col].median()),
            "Thiessen_area_max_m2": float(df[tp_area_col].max()),
            "Grid_weight_sum": grid_weight_sum,
            "CM_shift_m": math.hypot(grid_cm["x"] - CMx, grid_cm["y"] - CMy),
            **{f"CM_{k}_TMcal_m": v for k, v in (("x", CMx), ("y", CMy), ("as", CMas), ("cs", CMcs))},
            **{f"CM_{k}_Area_m": grid_cm[k] for k in ("x", "y", "as", "cs")},
            **{f"Variance_{k}_TMcal_m2": variances[k] for k in ("x", "y", "as", "cs")},
            **{f"Variance_{k}_Area_m2": grid_variances[k] for k in ("x", "y", "as", "cs")},
            "Variance_total_TMcal_m2": variances["x"] + variances["y"],
            "Variance_total_Area_m2": grid_variances["x"] + grid_variances["y"],
            "Variance_total_difference_pct": (
                100 * ((grid_variances["x"] + grid_variances["y"]) /
                       (variances["x"] + variances["y"]) - 1)
                if (variances["x"] + variances["y"]) > 0 else float("nan")
            ),
        })
        grid_interval_rows.append({
            "Sampling intervals": interval_label,
            "Interval_dt_s": delta_t_seconds,
            **{f"D_{k}_TMcal_1e-8": scaled_1e8(v) for k, v in (("x", D_x), ("y", D_y), ("as", D_as), ("cs", D_cs))},
            **{f"D_{k}_Area_1e-8": scaled_1e8(grid_D[k]) for k in ("x", "y", "as", "cs")},
            "D_total_TMcal_1e-8": scaled_1e8(D_x + D_y),
            "D_total_Area_1e-8": scaled_1e8(grid_D["x"] + grid_D["y"]),
        })
        prev_grid_variances = grid_variances if valid_grid else None
        prev_ip_moments = ip_moments if TM_sum > 0 else prev_ip_moments
        prev_variances = variances if TM_sum > 0 else None

        # Update previous_end_date, previous CM coordinates and label for the *next* sheet
        previous_end_date = this_end_date
        prev_CMx = CMx
        prev_CMy = CMy
        prev_CMas = CMas
        prev_CMcs = CMcs
        prev_label = sheet_name

    # -------------------------------------------------
    # 4. Build Vm/Q summary DataFrame (scaled numeric)
    # -------------------------------------------------
    vmq_df = pd.DataFrame(
        vmq_rows,
        columns=[
            "Sampling intervals",
            "Vmx", "Vmy", "Vmas", "Vmcs",
            "Qx", "Qy", "Qas", "Qcs",
            "Q_xy_resultant", "Q_as_cs_resultant",
        ]
    )

    # -------------------------------------------------
    # 5. Build variance-based diffusion summary DataFrame (Table 3) (scaled numeric)
    # -------------------------------------------------
    diff_df = pd.DataFrame(
    diff_rows,
    columns=[
        "Sampling intervals",
        "Dx", "Dy", "Das", "Dcs",
        "D_total",
    ]
    )

    # -------------------------------------------------
    # 5b. Build TCOM displacement summary DataFrame
    # -------------------------------------------------
    tcom_df = pd.DataFrame(
        tcom_rows,
        columns=[
            "Sampling intervals", "From", "To",
            "prev_EndDate", "End_date", "delta_t_s",
            "prev_CM_x", "prev_CM_y", "CM_x", "CM_y",
            "DistCM_x_m", "DistCM_y_m", "TCOMd_xy_m", "TCOMdir_deg_xy",
            "prev_CM_as", "prev_CM_cs", "CM_as", "CM_cs",
            "DistCM_as_m", "DistCM_cs_m", "TCOMd_ascs_m", "TCOMdir_deg_ascs",
        ]
    )

    # -------------------------------------------------
    # 6. Build Tracer Mass Recovery Summary Table
    # -------------------------------------------------
    recovery_rows = []

    # Ensure campaigns are in chronological order CM2, CM3, ... CM7
    campaign_order = sorted(
        com_sheets.keys(),
        key=lambda s: int(s[2:])  # CM2 -> 2, CM3 -> 3, etc.
    )

    for cm_name in campaign_order:
        cm_df = com_sheets[cm_name]
        row = cm_df.iloc[0]
        recovery_rows.append({
            "Campaign": cm_name,
            "T_rate_sum_kg": float(row["T_rate_sum"]),
            "T_rate_%": float(row["T_rate_%_sum"]),
        })

    # DataFrame with one row per CM campaign
    recovery_df = pd.DataFrame(recovery_rows)

    # Cumulative % recovery (per campaign)
    recovery_df["T_rate_%_cum"] = recovery_df["T_rate_%"].cumsum()

    # Add final TOTAL row
    total_T = recovery_df["T_rate_sum_kg"].sum()
    total_pct = recovery_df["T_rate_%"].sum()
    total_cum = recovery_df["T_rate_%_cum"].iloc[-1]

    recovery_df.loc[len(recovery_df.index)] = {
        "Campaign": "TOTAL",
        "T_rate_sum_kg": total_T,
        "T_rate_%": total_pct,
        "T_rate_%_cum": total_cum,
    }

    # -------------------------------------------------
    # 7. Write everything to the output workbook
    # -------------------------------------------------
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Write CM0 unchanged
        cM0.to_excel(writer, sheet_name="CM0", index=False)

        # Write processed C2..C7 sheets
        for name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

        # Write CM2..CM7 sheets (one row each)
        for cm_name, cm_df in com_sheets.items():
            cm_df.to_excel(writer, sheet_name=cm_name, index=False)

        # Write successive TCOM displacement summary table (unscaled, metres)
        tcom_df.to_excel(writer, sheet_name="TCOM_Displacements", index=False)

        # Write Vm/Q summary table (scaled ×10^-8, numeric)
        vmq_df.to_excel(writer, sheet_name="Vm_Q_table_10^(-8)", index=False)

        # Write Table 3 variance-based diffusion and Table 2 tracer spreading
        diff_df.to_excel(writer, sheet_name="Variance_Based_Diffusion", index=False)
        pd.DataFrame(spreading_rows).to_excel(writer, sheet_name="Tracer_Spreading_10^-8", index=False)
        pd.DataFrame(ip_dispersion_rows).to_excel(writer, sheet_name="IP_Referenced_Dispersion", index=False)
        pd.DataFrame(grid_sensitivity_rows).to_excel(writer, sheet_name="Sampling_Grid_Sensitivity", index=False)
        pd.DataFrame(grid_interval_rows).to_excel(writer, sheet_name="Grid_Diffusion_Comparison", index=False)

        # Write tracer mass recovery summary
        recovery_df.to_excel(
            writer,
            sheet_name="Tracer_Recovery_Summary",
            index=False
        )

    # -------------------------------------------------
    # 8. Apply Excel number format "0.00" to summary tables
    #     (keeps values numeric but always shows 2 decimals)
    # -------------------------------------------------
    wb = load_workbook(output_file)

    for sheet_name in [
        "Vm_Q_table_10^(-8)",
        "Variance_Based_Diffusion",
        "Tracer_Spreading_10^-8",
        "IP_Referenced_Dispersion",
        "Sampling_Grid_Sensitivity",
        "Grid_Diffusion_Comparison",
        "Tracer_Recovery_Summary",
        "TCOM_Displacements",
    ]:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Start from row 2 (skip header), col 2 (skip first text column)
            for row in ws.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "0.00"

    # -------------------------------------------------
    # 9. Generate tracer recovery histogram by campaign
    # -------------------------------------------------
    try:
        import matplotlib.pyplot as plt

        # Use recovery_df already computed (exclude the TOTAL row)
        recovery_plot_df = recovery_df[recovery_df["Campaign"] != "TOTAL"]

        fig, ax = plt.subplots()

        # Bar plot (histogram-like) of tracer recovery % per campaign

        campaign_labels = [
        c.replace("CM", "C")
        for c in recovery_plot_df["Campaign"]
        ]

        ax.bar(
        campaign_labels,
        recovery_plot_df["T_rate_%"],
        )

        ax.set_xlabel("Sampling campaign")
        ax.set_ylabel("Tracer recovery (%)")
        ax.set_title("Tracer mass recovery per campaign")
        fig.tight_layout()

        # Save as a new figure (rename if you prefer to overwrite the old one)
        fig.savefig("Tracer_Recovery_Histogram.png", dpi=300)
        plt.close(fig)
        print("📊 Saved plot: Tracer_Recovery_Histogram.png")

    except Exception as e:
        print(f"⚠ WARNING: Could not create recovery histogram: {e}")

    # Diagnostic plot: spatial variance (sum of Cartesian components)
    # in m², for both sampling-grid weighting schemes.
    try:
        import matplotlib.pyplot as plt
        grid_plot = pd.DataFrame(grid_sensitivity_rows)
        fig, ax = plt.subplots()
        ax.plot(grid_plot["Campaign"], grid_plot["Variance_total_TMcal_m2"],
                marker="o", label="TMcal (original)")
        ax.plot(grid_plot["Campaign"], grid_plot["Variance_total_Area_m2"],
                marker="s", label="TMcal × Thiessen area")
        ax.set_xlabel("Sampling campaign")
        ax.set_ylabel("Total central spatial variance (m²)")
        ax.set_title("Sampling-grid sensitivity of tracer spreading")
        ax.legend()
        fig.tight_layout()
        fig.savefig("Sampling_Grid_Sensitivity.png", dpi=300)
        plt.close(fig)
        print("📊 Saved plot: Sampling_Grid_Sensitivity.png")
    except Exception as e:
        print(f"⚠ WARNING: Could not create sampling-grid sensitivity plot: {e}")

    # -------------------------------------------------
    # 10. Save workbook
    # -------------------------------------------------
    wb.save(output_file)

    print("\n✅ Processing completed successfully.")
    print(f"📁 Output saved as: {output_file}")


if __name__ == "__main__":
    main()


