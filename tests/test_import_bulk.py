import ast
import unittest
from pathlib import Path
from contextlib import nullcontext
from types import SimpleNamespace
from psycopg2.extensions import adapt
from import_bulk import insert_detail_rows, DETAIL_INSERT_SQL


class Cursor:
    connection = SimpleNamespace(encoding='UTF8')
    def __init__(self): self.commands = []
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def mogrify(self, template, values):
        if isinstance(template, str): template = template.encode('utf-8')
        return template % tuple(adapt(value).getquoted() for value in values)
    def execute(self, sql): self.commands.append(sql)


class BulkImportTest(unittest.TestCase):
    def test_actual_insert_batches_and_quotes_values(self):
        # Extract the production INSERT so this tests the actual columns/casts/upsert.
        source = Path(__file__).resolve().parents[1].joinpath('app.py').read_text(encoding='utf-8')
        fn = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef) and n.name == 'save_policy_detail_to_postgres')
        assignment = next(n for n in ast.walk(fn) if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'insert_sql' for t in n.targets))
        sql = DETAIL_INSERT_SQL
        import re
        keys = re.findall(r':(\w+)', re.search(r'VALUES\s*\((.*?)\)\s*ON CONFLICT', sql, re.S).group(1))
        row = dict.fromkeys(keys, 1)
        row.update(source_file="O'Brien.xlsx", raw_data='{"name":"O\'Brien"}')
        cursor = Cursor()
        conn = SimpleNamespace(connection=SimpleNamespace(cursor=lambda: cursor))
        insert_detail_rows(conn, sql, [row.copy() for _ in range(4001)])
        self.assertEqual(len(cursor.commands), 3)
        self.assertIn(b"O''Brien.xlsx", cursor.commands[0])
        self.assertIn(b'CAST(', cursor.commands[0])
        self.assertIn(b'AS jsonb)', cursor.commands[0])
        self.assertIn(b'WHERE source_row_key IS NOT NULL DO UPDATE', cursor.commands[0])
        self.assertEqual(sum(command.count(b"O''Brien.xlsx") for command in cursor.commands), 4001)

    def test_database_failure_propagates_to_transaction(self):
        class FailedCursor(Cursor):
            def execute(self, sql): raise RuntimeError('database failure')
        conn = SimpleNamespace(connection=SimpleNamespace(cursor=lambda: FailedCursor()))
        with self.assertRaisesRegex(RuntimeError, 'database failure'):
            insert_detail_rows(conn, 'INSERT INTO example (a) VALUES (:a) ON CONFLICT (a) DO NOTHING', [{'a': 1}])


if __name__ == '__main__': unittest.main()
