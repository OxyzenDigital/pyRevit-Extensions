# Changelog

## [1.0.1] - 2025-12-22

### Changed
- **Collision Detection:** Removed Floors, Ceilings, and Roofs from collision checks (in addition to Walls). Pipes will now route through all architectural/host elements but will still avoid Structure and MEP services.

## [1.0.0] - 2025-12-22

### Added
- **Multi-Strategy Joining:** 
    - **Direct Bridge:** Attempts to connect pipes using a straight path (coplanar or skew closest points).
    - **Slide Bypass (Z-Shape):** Offsets the connection along the reference pipe's axis to avoid obstacles.
    - **Goal Post Bypass (U-Shape):** Attempts vertical (Up/Down) and horizontal (Side) jumps to route around obstacles.
- **Robust Collision Avoidance:**
    - Integrated multi-ray collision detection (Center + Perimeter) for pipe volume simulation.
    - Checks collision against Structure, Duct, Cable Tray, Conduit, and MEP Fittings.
    - **Wall Penetration:** Explicitly allows pipes to route through walls (Walls excluded from collision check).
    - **Extension Safety:** Verifies that extending existing pipes to the connection point does not cause new collisions.
- **Smart Connection Logic:**
    - **Auto-Fitting Selection:** Automatically chooses between `Elbow` (for angles) and `Union` (for collinear pipes).
    - **Fault Tolerance:** If fitting creation fails (e.g., due to missing families or strict routing preferences), the script **preserves the generated geometry** (pipes) instead of rolling back, allowing for manual fix-up.
    - **Parallel Pipe Handling:** Correctly projects the user's selection point to determine the best bridge location for parallel pipes.
- **Dynamic Offsets:** Bypass offsets are calculated based on pipe diameter (`max(4", 2.0 * Dia)`) to ensure adequate space for fittings.

### Fixed
- Fixed "failed to insert elbow" errors causing geometry to disappear by implementing robust error suppression and transaction handling.
- Fixed "too-short pipe" errors for parallel pipes by using user selection projection.
- Fixed issue where reference pipes extended through obstacles by adding collision checks to the extension logic.
