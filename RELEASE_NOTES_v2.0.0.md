# v2.0.0 Release Notes

This release introduces a major methodological update to the tracer metrics routine.

## Main Change

Tracer diffusion is now calculated using a method-of-moments formulation consistent with the tracer-spreading concept discussed by Miller and Komar. The routine now separates:

- transport/advection: centre-of-mass displacement between campaigns;
- diffusion/spreading: change in tracer-cloud variance between campaigns.

The diffusion coefficient is calculated as:

```text
D = delta_variance / (2 * delta_t)
```

where variance is TMcal-weighted and calculated around the centre of mass of each campaign.

## Why This Is v2.0.0

This update changes the scientific meaning of the diffusion output. Earlier versions mixed tracer-cloud displacement with tracer spreading. Version 2.0.0 treats diffusion as the time-rate of change of spatial variance, so diffusion values are not directly comparable with pre-v2.0.0 outputs.

## Additional Changes

- Added `Delta_Var_*` diagnostics to the diffusion summary.
- Added metadata describing the diffusion convention and first-interval assumption.
- Fixed automatic output naming so `.xlsx` files are no longer accidentally converted to `.xlsxx`.

## Interpretation Notes

For the first interval, the injection/reference variance is assumed to be zero. This means diffusion from injection to C1 is calculated from the variance observed in C1.

Negative directional diffusion values may occur when the variance in a given direction decreases between campaigns. These values should be retained and interpreted as evidence of tracer-cloud contraction, spatial reorganization, incomplete recovery, or sampling uncertainty in that directional component.
