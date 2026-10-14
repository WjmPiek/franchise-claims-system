import unittest
from flask import Flask
from period_selection import selected_dates


class LoginPeriodDefaultTest(unittest.TestCase):
    def setUp(self):
        self.app=Flask(__name__)
        self.app.secret_key='test-only'
        self.fallback={'date_from':'2024-08-01','date_to':'2026-09-30'}

    def test_fresh_session_opens_latest_available_month(self):
        with self.app.test_request_context('/'):
            self.assertEqual(selected_dates({},self.fallback),
                             {'date_from':'2026-09-01','date_to':'2026-09-30'})

    def test_selected_custom_dates_remain_across_pages(self):
        with self.app.test_request_context('/'):
            dates={'date_from':'2026-01-12','date_to':'2026-04-18'}
            self.assertEqual(selected_dates(dates,self.fallback),dates)
        with self.app.test_request_context('/'):
            from flask import session
            session['period_dates']=dates
            session['period_default_version']=2
            self.assertEqual(selected_dates({},self.fallback),dates)

    def test_latest_february_uses_calendar_month_end(self):
        with self.app.test_request_context('/'):
            self.assertEqual(selected_dates({}, {'date_from':'2020-01-01','date_to':'2024-02-29'}),
                             {'date_from':'2024-02-01','date_to':'2024-02-29'})

    def test_existing_old_session_resets_once(self):
        from flask import session
        with self.app.test_request_context('/'):
            session['period_dates']=self.fallback
            self.assertEqual(selected_dates({},self.fallback)['date_from'],'2026-09-01')
        with self.app.test_request_context('/'):
            custom={'date_from':'2026-01-01','date_to':'2026-04-30'}
            self.assertEqual(selected_dates(custom,self.fallback),custom)
            self.assertEqual(selected_dates({},self.fallback),custom)

    def test_login_reset_overrides_old_redirect_dates_once(self):
        from flask import session
        with self.app.test_request_context('/'):
            session['latest_period_pending']=True
            self.assertEqual(selected_dates(self.fallback,self.fallback)['date_from'],'2026-09-01')
            self.assertNotIn('latest_period_pending',session)
            self.assertEqual(selected_dates(self.fallback,self.fallback)['date_from'],'2026-09-01')
        with self.app.test_request_context('/'):
            self.assertEqual(selected_dates(self.fallback,self.fallback),self.fallback)

    def test_successful_login_sets_reset_flag(self):
        from flask import session,redirect
        from report_workflow import register
        register(self.app)
        @self.app.post('/login')
        def login():
            session['user_id']=1
            session['period_dates']=self.fallback
            return redirect('/overview/risk?date_from=2025-10-01&date_to=2026-09-30')
        client=self.app.test_client()
        self.assertEqual(client.post('/login').status_code,302)
        with client.session_transaction() as saved:
            self.assertTrue(saved['latest_period_pending'])
            self.assertNotIn('period_dates',saved)


if __name__=='__main__':
    unittest.main()
