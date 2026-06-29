# Gorden detail pages todo

The final all-factor animation must not use programmatic detail layers.
Each detail page must first be rebuilt through GordenImage2PPTX from the
high-quality static reference image.

## Required input images

- `F:\桌面\Keynote缩略图\因素矩阵\static_reference_states_20260627\01_temperature_f01.png`
- `F:\桌面\Keynote缩略图\因素矩阵\static_reference_states_20260627\02_pressure_f02.png`
- `F:\桌面\Keynote缩略图\因素矩阵\static_reference_states_20260627\03_material_f03_source.png`
- `F:\桌面\Keynote缩略图\因素矩阵\static_reference_states_20260627\04_roughness_f04.png`
- `F:\桌面\Keynote缩略图\因素矩阵\static_reference_states_20260627\05_contact_area_f05.png`
- `F:\桌面\Keynote缩略图\因素矩阵\static_reference_states_20260627\06_ice_composition_f06.png`

## Required output directory

`F:\桌面\Keynote缩略图\因素矩阵\gorden_detail_pages_strict`

## Required PPTX names

- `gorden_detail_01_temperature_f01.pptx`
- `gorden_detail_02_pressure_f02.pptx`
- `gorden_detail_03_material_f03.pptx`
- `gorden_detail_04_roughness_f04.pptx`
- `gorden_detail_05_contact_area_f05.pptx`
- `gorden_detail_06_ice_composition_f06.pptx`

## Acceptance gate

Every page directory must include `imagegen-assets-manifest.json` with
real imagegen evidence for background, frame, and icon/decor layers.

The animation integration script intentionally refuses:

- `gorden_static_single_pages`
- `GordenImage2PPTX-inspired` outputs
- programmatic/local/PIL/SVG/HTML/Canvas/matplotlib image-layer sources

After the strict Gorden pages exist, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_factor_matrix_all_animations_from_gorden_pages.ps1
```
