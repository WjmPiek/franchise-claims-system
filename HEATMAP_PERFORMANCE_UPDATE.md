# Client heatmap performance update

The client heatmap now loads prepared coordinates by viewport and zoom instead
of downloading every address group in sequential 500-row pages.

- Country view returns weighted density cells.
- Pan and zoom requests are debounced.
- Responses and rendered overlays are capped at 1,500 by default.
- Selected franchises reveal address-level pins at zoom 12 and closer.
- Country-wide responses omit client names, policy numbers, and addresses.
- The legacy paged API remains available for compatibility.

`heatmap_performance_migration.sql` contains an optional concurrent production
index for faster selected-franchise viewport queries.
