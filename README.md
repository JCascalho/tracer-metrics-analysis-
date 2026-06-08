# Tracer Metrics Analysis

Python routine for calculating sediment-tracer recovery, centre-of-mass displacement, diffusion, transport velocity, and transport-flux metrics from a multi-campaign tracer-monitoring workbook.

The routine reads an Excel workbook with an injection/reference sheet and campaign sheets, projects Cartesian `x`/`y` offsets into along- and cross-sand-body components, and exports processed campaign tables and summary metrics.

## Features

- Processes tracer-monitoring workbooks with `IP` and campaign sheets such as `C1` to `C5`.
- Calculates calibrated tracer mass concentration (`TMcal`).
- Calculates centre-of-mass displacement between campaigns.
- Projects displacement vectors into along- and cross-sand-body components.
- Calculates horizontal diffusion metrics.
- Calculates mean tracer transport velocities and transport fluxes.
- Calculates tracer-grain density where photo-counting columns are available.
- Calculates tracer-mass recovery and cumulative recovery.
- Exports a structured Excel workbook with campaign outputs, centre-of-mass tables, summary metrics, and metadata.

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Input Workbook

The workbook must contain:

- `IP`: injection/reference sheet
- one or more campaign sheets, for example `C1`, `C2`, `C3`, `C4`, `C5`

Minimum required columns:

- `SAMPLE_ID`
- `DATE_TIME`
- `x`
- `y`
- `Ac`
- `m`
- `Tp_area`

Recommended optional columns:

- `tracer_nr`
- `photo_area`
- `dreger_d`
- `sed_poro`
- `rho_sed`
- `tracer_m`
- `delta_mix`

## Usage

Basic run:

```bash
python tracer_metrics_analysis.py --input INPUT_TRACER_DATA.xlsx
```

Specify output file:

```bash
python tracer_metrics_analysis.py --input INPUT_TRACER_DATA.xlsx --output OUTPUT_TRACER_METRICS.xlsx
```

Process selected campaign sheets:

```bash
python tracer_metrics_analysis.py --input INPUT_TRACER_DATA.xlsx --campaigns C1,C2,C3
```

Change sand-body azimuths:

```bash
python tracer_metrics_analysis.py --input INPUT_TRACER_DATA.xlsx --along-azimuth 41 --cross-azimuth 131
```

Spyder example:

```python
%runfile "C:/path/to/tracer_metrics_analysis.py" --wdir "C:/path/to/tracer-metrics-analysis" --args "--input INPUT_TRACER_DATA.xlsx --output OUTPUT_TRACER_METRICS.xlsx"
```

## Outputs

The output workbook includes:

- processed campaign sheets
- centre-of-mass sheets
- `Velocity_Flux_1e-8`
- `Diffusion_1e-8`
- `Tracer_Recovery`
- `Metadata`

## Notes

Along- and cross-sand-body components are computed from Cartesian x/y offsets using the selected azimuths. Existing `asb`/`csb` columns in the input workbook are not used for diffusion, velocity, or flux calculations.

## Citation

If you use this routine, please cite the archived GitHub release DOI generated through Zenodo.
