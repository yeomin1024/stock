# CHANGELOG — seoul-modular-building

## v1.0.0 — 2026-11-02 — Initial implementation

**Added**
- Vite 8 + TypeScript 5.9 + three.js r186 project; procedural model built only from a shared `BoxGeometry(1,1,1)`.
- `src/config.ts`: single source of truth for every dimension (envelope, floors, bay, facade, entrance, canopy, parapet, roof equipment), materials, lighting and UI.
- `src/derive.ts`: floor levels, glazing rows, solid bands and entrance zone derived from the config (no hand-typed levels).
- `src/layout.ts`: pure box layout (ids + tags on every box) for the 8 groups under `BuildingRoot`.
- `src/validate.ts`: config checks (height sum, bay tiling, entrance fit, reveal depth, equipment inside parapet) and layout audit (18.0 × 12.0 × 24.0 envelope, 2.4 × 2.4 module, entrance size, glass/opaque intersections, intentional opaque overlaps).
- `src/building/*`: `InstancedMesh` per repeated part; glass in separate transparent meshes; per-group triangle/draw-call stats.
- `src/stage.ts`: renderer (NeutralToneMapping, sRGB output, PCF soft shadows), hemisphere + directional sun, RoomEnvironment IBL, ground with shadows, toggleable 3.0 m `GridHelper`, exact camera framing, view presets.
- `src/demo.ts`: deterministic assembly + turntable timeline (`seek(t)`), used by the UI and by the video capture.
- `src/logger.ts`: structured `[STAGE] [timestamp] key=value` logging, level via `?log=debug`.
- `tests/layout.test.ts`: 17 unit tests (dimensions, counts, module sizes, corner condition, overlaps, guard checks).
- `scripts/capture.mjs`: optional headless stills + frame-exact MP4 capture.

**Fixed during verification (before release)**
- `layout.ts` / `derive.ts` — corner geometry: windows centered in 3.0 m bays leave 0.30 m corner frames, but frames are 0.35 m deep, so a glass plane at the full 0.35 m depth made the two corner panes of adjacent faces cross each other's reveal (48 glass/mullion intersections caught by `GLASS_OPAQUE_OVERLAP`). The frames keep their 0.35 m depth; the glass plane is now set back `min(cornerWidth, projectionDepth)` = 0.30 m so corner panes meet exactly at the corner pier. Covered by a dedicated test.
- `instancing.ts` — glass now casts shadows: non-casting glass let sunlight through both facades and speckled the ground shadow with window-shaped light patches.
- `config.ts` — material/light tuning after visual review (frames rendered near-black, glass too clear, concrete too warm).
- `demo.ts` — camera stays in the front-right quadrant during assembly (entrance/canopy steps were happening while the camera faced the rear), then does the full turntable; timeline 14 s → 16 s.
