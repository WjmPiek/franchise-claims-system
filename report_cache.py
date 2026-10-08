"""Bounded report caches keyed by committed imports and mapping content."""
from collections import OrderedDict
import threading
import time

class ResultCache:
    def __init__(self, ttl=30, max_entries=24):
        self.ttl, self.max_entries = ttl, max_entries
        self._values = OrderedDict()
        self._lock = threading.RLock()

    def get(self, key, loader):
        # Serialize identical generation and bound the number of in-memory results.
        with self._lock:
            current = self._values.get(key)
            if current and time.monotonic() - current[0] < self.ttl:
                self._values.move_to_end(key)
                return current[1]
            value = loader()  # Failures never become cached successes.
            self._values[key] = (time.monotonic(), value)
            self._values.move_to_end(key)
            while len(self._values) > self.max_entries:
                self._values.popitem(last=False)
            return value

    def clear(self):
        with self._lock:
            self._values.clear()

DETAIL_CACHE = ResultCache(ttl=30)
REFRESH_CACHE = ResultCache(ttl=30, max_entries=4)

def data_revision(connection, sql_text):
    row = next(iter(connection.execute(sql_text("""
        SELECT COALESCE((SELECT MAX(id) FROM import_history), 0) AS imports,
               COALESCE((SELECT md5(string_agg(
                   COALESCE(source_name,'') || '|' || COALESCE(mapped_name,'') || '|' ||
                   COALESCE(approved,false)::text, ';' ORDER BY id))
                   FROM franchise_mapping_pg), '') AS mapping
    """), {}).mappings()))
    return (row['imports'], row['mapping'])

def refresh_reporting_data(engine, sql_text, refresh, settings=()):
    with engine.connect() as connection:
        revision = data_revision(connection, sql_text)
    return REFRESH_CACHE.get((engine, revision, settings), refresh)

def read_policy_details(engine, sql_text, query, params):
    with engine.connect() as connection:
        revision = data_revision(connection, sql_text)
    key = (engine, revision, query, tuple(sorted(params.items())))
    def read():
        with engine.connect() as connection:
            return tuple(dict(row) for row in connection.execute(sql_text(query), params).mappings())
    return DETAIL_CACHE.get(key, read)
