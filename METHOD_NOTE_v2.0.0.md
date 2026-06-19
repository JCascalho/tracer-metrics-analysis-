# Method Note - Tracer Metrics v2.0.0

Version 2.0.0 revises the diffusion calculation to better match the physical concept of tracer spreading.

## Transport

Transport velocity is calculated from the displacement of the tracer centre of mass between consecutive monitoring campaigns:

```text
Vm_x = (CM_x_current - CM_x_previous) / delta_t
Vm_y = (CM_y_current - CM_y_previous) / delta_t
```

The same vector is projected onto the along-sandbank and cross-sandbank axes:

```text
Vm_asb, Vm_csb
```

Flux proxies are then calculated from the velocity components and the active mixing thickness.

## Diffusion

Diffusion is calculated from the time-rate of change of weighted spatial variance. For each campaign:

```text
Var_x   = sum(TMcal * (x   - CM_x)^2)   / sum(TMcal)
Var_y   = sum(TMcal * (y   - CM_y)^2)   / sum(TMcal)
Var_asb = sum(TMcal * (asb - CM_asb)^2) / sum(TMcal)
Var_csb = sum(TMcal * (csb - CM_csb)^2) / sum(TMcal)
```

For each monitoring interval:

```text
D_x   = (Var_x_current   - Var_x_previous)   / (2 * delta_t)
D_y   = (Var_y_current   - Var_y_previous)   / (2 * delta_t)
D_asb = (Var_asb_current - Var_asb_previous) / (2 * delta_t)
D_csb = (Var_csb_current - Var_csb_previous) / (2 * delta_t)
```

The first interval assumes:

```text
Var_x_previous = Var_y_previous = Var_asb_previous = Var_csb_previous = 0
```

This represents a compact tracer release at the injection/reference point.

## Physical Meaning

The revised method separates advection from spreading:

- centre-of-mass displacement describes net tracer transport;
- variance growth describes tracer-cloud spreading/diffusion.

This distinction is important in field settings where wave-current forcing may move a compact tracer patch without necessarily producing the largest diffusion coefficient.
