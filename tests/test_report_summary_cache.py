import unittest
import os
import time
import threading
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import Mock, patch
from report_cache import ResultCache, read_policy_details, refresh_reporting_data, DETAIL_CACHE, REFRESH_CACHE

class ReportCacheTest(unittest.TestCase):
    def setUp(self):
        DETAIL_CACHE.clear()
        REFRESH_CACHE.clear()
        self.revision = 1
        self.mapping = ''
        self.queries = 0
        outer = self
        class Result:
            def __init__(self, rows): self.rows = rows
            def mappings(self): return self.rows
        class Connection:
            def __enter__(self): return self
            def __exit__(self,*args): pass
            def execute(self,query,params):
                if 'FROM import_history' in query:
                    return Result([{'imports':outer.revision,'mapping':outer.mapping}])
                outer.queries += 1
                return Result([{'amount':outer.revision}])
        class Engine:
            def connect(self): return Connection()
        self.engine = Engine()

    def test_detail_scan_reused_between_formats_but_import_invalidates(self):
        for unused in range(4):
            self.assertEqual(read_policy_details(self.engine,lambda s:s,'DETAIL',{'month':'2024-09'})[0]['amount'],1)
        self.assertEqual(self.queries,1)
        self.revision = 2
        self.assertEqual(read_policy_details(self.engine,lambda s:s,'DETAIL',{'month':'2024-09'})[0]['amount'],2)
        self.assertEqual(self.queries,2)
        read_policy_details(self.engine,lambda s:s,'DETAIL',{'month':'2024-10'})
        self.assertEqual(self.queries,3)
        read_policy_details(self.engine,lambda s:s,'DETAIL',{'month':'2024-09','franchise':'Other user'})
        self.assertEqual(self.queries,4)

    def test_snapshot_refresh_reused_and_settings_mapping_imports_invalidate(self):
        loader = Mock(return_value=True)
        for unused in range(3):
            refresh_reporting_data(self.engine,lambda s:s,loader,('rates',))
        self.assertEqual(loader.call_count,1)
        self.revision += 1
        refresh_reporting_data(self.engine,lambda s:s,loader,('rates',))
        self.mapping = 'changed mapping'
        refresh_reporting_data(self.engine,lambda s:s,loader,('rates',))
        refresh_reporting_data(self.engine,lambda s:s,loader,('different rates',))
        self.assertEqual(loader.call_count,4)

    def test_different_scopes_are_not_serialized(self):
        cache = ResultCache()
        barrier = threading.Barrier(2)
        def load():
            barrier.wait(timeout=2)
            return 'ready'
        with ThreadPoolExecutor(max_workers=2) as pool:
            a = pool.submit(cache.get, 'September', load)
            b = pool.submit(cache.get, 'October', load)
            self.assertEqual(a.result(), 'ready')
            self.assertEqual(b.result(), 'ready')

    def test_expiry_failures_bounded_size_and_single_generation(self):
        cache = ResultCache(ttl=1,max_entries=2)
        loader = Mock(return_value=7)
        with patch('report_cache.time.monotonic',return_value=0):
            cache.get('key',loader)
        with patch('report_cache.time.monotonic',return_value=2):
            cache.get('key',loader)
        self.assertEqual(loader.call_count,2)
        cache.get('two',loader)
        cache.get('three',loader)
        self.assertEqual(len(cache._values),2)
        failure = Mock(side_effect=RuntimeError('database unavailable'))
        for unused in range(2):
            with self.assertRaises(RuntimeError):
                cache.get('bad',failure)
        self.assertEqual(failure.call_count,2)
        count = []
        def slow():
            count.append(1)
            time.sleep(.03)
            return 123
        with ThreadPoolExecutor(max_workers=4) as pool:
            results = list(pool.map(lambda _:cache.get('same',slow),range(4)))
        self.assertEqual(results,[123]*4)
        self.assertEqual(len(count),1)

@unittest.skipUnless(os.environ.get('PRECISION_TEST_DATABASE_URL'), 'requires isolated PostgreSQL')
class RevisionSqlTest(unittest.TestCase):
    def test_actual_postgres_import_and_mapping_revision(self):
        from sqlalchemy import create_engine, text
        from report_cache import data_revision
        engine = create_engine(os.environ['PRECISION_TEST_DATABASE_URL'])
        try:
            with engine.begin() as conn:
                conn.execute(text('CREATE TABLE IF NOT EXISTS import_history (id BIGSERIAL PRIMARY KEY)'))
                conn.execute(text('CREATE TABLE IF NOT EXISTS franchise_mapping_pg (id BIGSERIAL PRIMARY KEY, source_name TEXT, mapped_name TEXT, approved BOOLEAN)'))
                before = data_revision(conn,text)
                conn.execute(text('INSERT INTO import_history DEFAULT VALUES'))
                after = data_revision(conn,text)
                self.assertGreater(after[0],before[0])
                conn.execute(text("INSERT INTO franchise_mapping_pg (source_name,mapped_name,approved) VALUES ('A','B',true)"))
                mapped = data_revision(conn,text)
                self.assertNotEqual(after[1],mapped[1])
        finally:
            engine.dispose()

if __name__ == '__main__':
    unittest.main()
