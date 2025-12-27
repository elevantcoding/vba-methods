# security
Cross-platform encryption system (SQL / VBA / Python)

This folder contains cross-plataform encryption utilities designed to operate consistently across:
- ** SQL Server (T-SQL)**
- ** VBA / Microsft Access**
- ** Python**
- The goal of this design is deterministic compatibility
- data encrypted in one environment can be decrypted correctly in the others.
- ## Highlights
- Custom cipher with randomized components
- Deterministic encryption / decryption across platforms
- No dependency on external crypto libraries
- Safe for use in mixed-technology stacks (Access - SQL Server - Python)

- ## Cross-Platform Compatibility
- The same encrypted value produced by SQL Server can be:
- decrypted by VBA
- decrypted by Python
- and verified again by SQL Server

- This allows encrypted values to move freely between systems without data loss.

- ## Contents
- Each language repository contains it's own implementations of the same algorithm.
