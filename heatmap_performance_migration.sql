-- Optional production index for the viewport-based client heatmap.
-- Run this statement by itself in DBeaver; do not wrap it in BEGIN/COMMIT.
CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_app_client_map_points_franchise_latlng
ON app_client_map_points (franchise_key, lat, lng)
WHERE lat IS NOT NULL AND lng IS NOT NULL;
