import ast
from datetime import timedelta
from pathlib import Path
import unittest
from unittest.mock import patch
from flask import Flask, g, session, request, jsonify, redirect, url_for, flash


class ImportSessionTest(unittest.TestCase):
    def setUp(self):
        self.app = Flask(__name__)
        self.app.secret_key = 'session-test'
        self.app.permanent_session_lifetime = timedelta(minutes=30)
        self.user = {'id': 1, 'is_active': True, 'role': 'admin'}
        tree = ast.parse(Path(__file__).resolve().parents[1].joinpath('app.py').read_text(encoding='utf-8'))
        nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef)
                 and n.name in {'load_logged_in_user', 'import_session'}]
        for node in nodes:
            node.decorator_list = []
        env = dict(g=g, session=session, request=request, jsonify=jsonify,
                   redirect=redirect, url_for=url_for, flash=flash,
                   get_user_by_id=lambda _: self.user, MAINTENANCE_MODE=False,
                   PUBLIC_ENDPOINTS={'login'}, ADMIN_ENDPOINTS=set(), DATABASE_URL='',
                   record_user_activity=lambda _: None)
        exec(compile(ast.Module(body=nodes, type_ignores=[]), 'auth-functions', 'exec'), env)
        self.app.before_request(env['load_logged_in_user'])
        self.app.add_url_rule('/api/import/session', 'import_session', env['import_session'])
        self.app.add_url_rule('/dashboard', 'dashboard', lambda: jsonify(ok=True), methods=['GET', 'POST'])
        self.app.add_url_rule('/login', 'login', lambda: 'login')
        self.client = self.app.test_client()

    def login(self):
        with self.client.session_transaction() as state:
            state['user_id'] = 1
            state.permanent = True

    def test_valid_session_renewed_without_extending_idle_timeout(self):
        self.login()
        response = self.client.get('/api/import/session')
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json['ok'])
        self.assertIn('session=', response.headers['Set-Cookie'])
        self.assertEqual(response.headers['Cache-Control'], 'no-store')
        self.assertEqual(self.app.permanent_session_lifetime, timedelta(minutes=30))

    def test_unauthenticated_import_returns_json_and_normal_page_redirects(self):
        for response in [self.client.get('/api/import/session'),
                         self.client.post('/dashboard', headers={'X-Requested-With': 'XMLHttpRequest'})]:
            self.assertEqual(response.status_code, 401)
            self.assertFalse(response.json['ok'])
            self.assertIn('session expired', response.json['error'])
        self.assertEqual(self.client.get('/dashboard').status_code, 302)

    def test_expired_cookie_cannot_be_revived_by_keepalive(self):
        with patch('itsdangerous.timed.TimestampSigner.get_timestamp', return_value=1000):
            self.login()
        with patch('itsdangerous.timed.TimestampSigner.get_timestamp', return_value=3000):
            response = self.client.get('/api/import/session')
        self.assertEqual(response.status_code, 401)

    def test_disabled_and_viewer_accounts_cannot_import(self):
        self.login()
        self.user['role'] = 'viewer'
        self.assertEqual(self.client.get('/api/import/session').status_code, 403)
        self.assertEqual(self.client.post('/dashboard', headers={'X-Requested-With': 'XMLHttpRequest'}).status_code, 403)
        self.user['is_active'] = False
        self.assertEqual(self.client.get('/api/import/session').status_code, 403)


if __name__ == '__main__':
    unittest.main()
