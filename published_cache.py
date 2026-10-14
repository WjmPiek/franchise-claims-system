"""Read committed publication versions before reusing derived calculations."""
from flask import has_request_context,g
from report_cache import ResultCache
FINANCIAL=ResultCache(ttl=3600,max_entries=4)
CLAIM_SOURCE=ResultCache(ttl=3600,max_entries=4)

def signature(store):
    key=(id(store),str(store.direct.path))
    if has_request_context() and getattr(g,'published_signature_key',None)==key:return g.published_signature
    store.ensure();d=store.direct
    with d.connection() as conn:
        policies=tuple(tuple(r) for r in d.execute(conn,'SELECT month,batch_id,version,published_at FROM central_publications ORDER BY month'))
        claims=tuple(tuple(r) for r in d.execute(conn,'SELECT p.month,p.batch_id,e.version FROM central_claim_publications p LEFT JOIN central_claim_event_versions e ON e.batch_id=p.batch_id ORDER BY p.month'))
        exists=d.execute(conn,"SELECT to_regclass('public.central_allocation_audit')").fetchone()[0] if d.get_engine else conn.execute("SELECT name FROM sqlite_master WHERE name='central_allocation_audit'").fetchone()
        audit=d.execute(conn,'SELECT MAX(changed_at) FROM central_allocation_audit').fetchone()[0] if exists else None
    value=(key,policies,claims,audit)
    if has_request_context():g.published_signature_key=key;g.published_signature=value
    return value
