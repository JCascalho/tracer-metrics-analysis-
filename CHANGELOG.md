# Changelog

## v2.0.0 - 2026-06-19

### Changed
- Reworked tracer diffusion calculations to follow a Miller and Komar-style method-of-moments approach.
- Diffusion is now calculated from the time-rate of change of TMcal-weighted tracer-cloud variance:
  `D = delta_variance / (2 * delta_t)`.
- Transport velocity and flux remain based on centre-of-mass displacement between monitoring campaigns.
- Added variance diagnostics to the diffusion output, including `Delta_Var_x`, `Delta_Var_y`, `Delta_Var_asb`, and `Delta_Var_csb`.
- Added metadata fields documenting the diffusion method, the first-interval variance assumption, and the transport-distance mode.

### Fixed
- Fixed automatic output filename generation to prevent invalid `.xlsxx` output files.

### Notes
- This is a methodological breaking change. Diffusion values from v2.0.0 are not directly comparable with pre-v2.0.0 results.
- The first interval assumes the injection/reference variance is zero, consistent with a compact release-point approximation.
- Negative directional diffusion components can occur when variance decreases between campaigns; these values should be interpreted as tracer-cloud contraction, redistribution, or sampling uncertainty in that direction rather than as literal negative physical diffusion.
