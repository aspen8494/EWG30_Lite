# EWG3.0 Formulator Lite

Vendored EWG3.0 formulation engine, byte-copied from the full application
(repo `HWG_mission_planning`, commit recorded in `VENDORED_SOURCE_COMMIT.txt`
at the repo root) so the lite CLI solves the identical optimization problem.
`gfc.py` / `properties.py` load their coefficient JSONs from
`data/ewg30/` relative to module location — self-contained.
