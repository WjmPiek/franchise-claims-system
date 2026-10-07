import io
import json
import unittest
from flask import Flask, jsonify, request
from import_progress import active, report, stream_import


class StreamImportTest(unittest.TestCase):
    def make_app(self, fail=False):
        app = Flask(__name__)
        app.secret_key = 'test'
        @app.post('/import')
        def view():
            if not active(): return stream_import(view)
            self.assertEqual(request.files['file'].read(), b'example workbook')
            report('Reading Excel rows', 2, 4)
            report('Writing policy rows (awaiting commit)', 4, 4)
            if fail: raise ValueError('example failure')
            return jsonify(ok=True)
        return app

    def test_multipart_context_counts_and_completion(self):
        response = self.make_app().test_client().post('/import', data={'file':(io.BytesIO(b'example workbook'),'test.xlsx')})
        updates = [json.loads(line) for line in response.data.splitlines()]
        self.assertEqual(response.mimetype, 'application/x-ndjson')
        self.assertEqual(updates[0]['current'], 2)
        self.assertEqual(updates[0]['total'], 4)
        self.assertEqual(updates[-1]['type'], 'done')
        self.assertTrue(updates[-1]['ok'])
        self.assertFalse(active())

    def test_failure_is_explicit_and_never_complete(self):
        response = self.make_app(True).test_client().post('/import', data={'file':(io.BytesIO(b'example workbook'),'test.xlsx')})
        updates = [json.loads(line) for line in response.data.splitlines()]
        self.assertFalse(updates[-1]['ok'])
        self.assertEqual(updates[-1]['type'], 'done')


if __name__ == '__main__': unittest.main()
