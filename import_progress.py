"""Stream real import stage/count updates over the existing import request."""
import json
import queue
import threading
import time
from flask import Response, copy_current_request_context, current_app, stream_with_context

_state = threading.local()
_import_lock = threading.Lock()


def busy():
    return _import_lock.locked()


def active():
    return bool(getattr(_state, 'emit', None))


def report(stage, current=None, total=None):
    emit = getattr(_state, 'emit', None)
    if emit:
        emit({'type': 'progress', 'stage': stage, 'current': current, 'total': total})


def stream_import(view):
    updates = queue.Queue()
    app = current_app._get_current_object()

    @copy_current_request_context
    def run():
        if not _import_lock.acquire(blocking=False):
            updates.put({'type': 'done', 'ok': False, 'error': 'An import is already running. Wait for it to finish before retrying.'})
            return
        _state.emit = updates.put
        try:
            response = app.make_response(view())
            result = response.get_json(silent=True) or {}
            updates.put({'type': 'done', 'ok': response.status_code < 300 and result.get('ok') is True,
                         'error': result.get('error', 'The server did not confirm a successful import.')})
        except Exception as exc:
            app.logger.error('Streamed Excel import failed (%s)', type(exc).__name__)
            updates.put({'type': 'done', 'ok': False, 'error': 'Import failed. Check this month before retrying.'})
        finally:
            del _state.emit
            _import_lock.release()

    @stream_with_context
    def events():
        started = time.monotonic()
        worker = threading.Thread(target=run, daemon=True)
        worker.start()
        while True:
            try:
                update = updates.get(timeout=5)
            except queue.Empty:
                update = {'type': 'heartbeat'}
            update['elapsed'] = round(time.monotonic() - started)
            yield json.dumps(update) + '\n'
            if update['type'] == 'done':
                break

    return Response(events(), mimetype='application/x-ndjson', headers={
        'Cache-Control': 'no-cache, no-transform', 'X-Accel-Buffering': 'no'})
