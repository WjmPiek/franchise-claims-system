"""Fast, parameterized PostgreSQL writes using the caller's transaction."""
import re
from psycopg2.extras import execute_values


def insert_detail_rows(connection, statement, rows):
    sql = str(statement)
    values = re.search(r'VALUES\s*\((.*?)\)\s*(ON CONFLICT)', sql, re.S)
    if values is None:
        raise ValueError('Expected detail INSERT with conflict handling')
    keys = re.findall(r':(\w+)', values.group(1))
    template = re.sub(r':\w+', '%s', values.group(1))
    bulk_sql = sql[:values.start()] + 'VALUES %s ' + sql[values.start(2):]
    # psycopg2 adapts every value; filenames, names and JSON never become SQL.
    # Keep the delete and every insert in the same SQLAlchemy transaction.
    with connection.connection.cursor() as cursor:
        for start in range(0, len(rows), 2000):
            batch = [tuple(row[key] for key in keys) for row in rows[start:start + 2000]]
            execute_values(cursor, bulk_sql, batch, template='(' + template + ')', page_size=2000)
