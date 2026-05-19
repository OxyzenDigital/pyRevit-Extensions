# Changelog - Grading Assistant

## [v1.0] - 2025-02-04
### Changed
- **Refactoring:**
  - Converted comment-based metadata to standard python variables (`__title__`, `__doc__`, etc.).
  - Cleaned up imports.

## [v1.1] - 2025-02-11
### Added
- **Precision Raycasting:** Integrated `ReferenceIntersector` raycasting for accurate halo elevation blending, guarded by a safety tolerance check against spatial averages.
- **Corner Offsets:** Added 2D vector intersection math to cleanly resolve mitered outside and inside corners for grading paths.
- **UI Integration:** Wired up state persistence, validation, and unit conversions for the `Point Dist Tol` setting.

### Changed / Refactored
- **Code Modularity:** Abstracted the Laplacian Smoothing engine (`apply_laplacian_smoothing`), Dual-Ring Halo Triangulation (`apply_triangulation_halo`), and region intersection checks (`get_region_z`) into reusable, globally accessible functions.
- **Smoothing Uniformity:** Both Region and Path smoothing now utilize the exact same optimized Laplacian engine.

### Optimized
- **Virtual Vertex Tracker:** Replaced heavy C# interop calls to `SlabShapeEditor.SlabShapeVertices` with a centralized, pure-Python dictionary tracker (`VirtualVertexTracker`), drastically reducing execution time across all grading functions.
- **Coarse Spatial Hashing:** Added a secondary spatial hash specifically tailored for wide-radius halo searches, eliminating thousands of redundant dictionary lookups.

### Fixed
- **Corrupted Vertices:** Added robust error swallowing in the `VirtualVertexTracker` to gracefully skip ghost/corrupted mesh vertices without crashing the script.
- **Tuple Unpacking:** Fixed `TypeError: 'SlabShapeVertex' object is not iterable` by standardizing vertex caching formats.
- **Syntax Strictness:** Resolved IronPython `IndentationError` crashes caused by inline `try/except` suites.