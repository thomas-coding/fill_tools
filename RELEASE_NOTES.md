# Release Notes

## v0.2 (2026-03-07)

### Highlights
- Added dual workflows: `巡检填报1` and `巡检填报2` with independent field mapping and operation tips.
- Separated status tracking by profile:
  - 填报1 uses `截止时间` 后第 1/2 列（状态/提交时间）
  - 填报2 uses `截止时间` 后第 3/4 列（状态/提交时间）
- Enhanced photo compatibility for both standard Excel and WPS-exported workbooks:
  - Standard drawing anchors (`xl/drawings`)
  - WPS `DISPIMG + cellimages`
  - Zip drawing/media fallback
- Added runtime diagnostics for missing photo rows to help locate `F12` path-empty issues.
- Bumped app version to `V0.2`.

### Notes
- Existing `F5/F12/F10` workflow remains unchanged.
- For `巡检填报2`, `F5` fills two text fields in order: `处置路段` then `整改描述`.
