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
            self.assertEqual(selected_dates({},self.fallback),dates)

    def test_latest_february_uses_calendar_month_end(self):
        with self.app.test_request_context('/'):
            self.assertEqual(selected_dates({}, {'date_from':'2020-01-01','date_to':'2024-02-29'}),
                             {'date_from':'2024-02-01','date_to':'2024-02-29'})


if __name__=='__main__':
    unittest.main()
