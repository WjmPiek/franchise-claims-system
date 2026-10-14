import unittest
import pandas as pd
from premium_comparison import comparison

def build(source,current,selected):
    start=current.Month.min()-pd.DateOffset(years=1)
    previous=source[source.Month==start]
    a=current['Retail Premium'].sum();b=previous['Retail Premium'].sum()
    return {'comparison_rows':[dict(current=a,previous=b,difference=a-b)],'has_previous':not previous.empty}

class PremiumComparisonTest(unittest.TestCase):
    def test_earlier_import_is_available_and_new_month_not_duplicated(self):
        current=pd.DataFrame([{'Franchise':'A','Month':'2026-09-01','Retail Premium':100}])
        old=pd.DataFrame([{'Franchise':'A','Month':'2025-09-01','Retail Premium':80},{'Franchise':'A','Month':'2026-09-01','Retail Premium':999}])
        result=comparison(current,old,'2026-09','2026-09','All',build)
        self.assertEqual(result['comparison_rows'][0],dict(current=100,previous=80,difference=20))
        self.assertEqual(result['historical_previous_months'],['2025-09'])
    def test_missing_month_does_not_become_zero(self):
        current=pd.DataFrame([{'Franchise':'A','Month':'2026-09-01','Retail Premium':100}])
        result=comparison(current,current.iloc[:0],'2026-09','2026-09','All',build)
        self.assertIsNone(result['comparison_rows'][0]['previous'])
        self.assertEqual(result['missing_previous_months'],['2025-09'])
    def test_franchise_selection_excludes_other_branches(self):
        current=pd.DataFrame([{'Franchise':'A','Month':'2026-09-01','Retail Premium':100}])
        old=pd.DataFrame([{'Franchise':'A','Month':'2025-09-01','Retail Premium':80},{'Franchise':'B','Month':'2025-09-01','Retail Premium':500}])
        self.assertEqual(comparison(current,old,'2026-09','2026-09','A',build)['comparison_rows'][0]['previous'],80)

if __name__=='__main__':unittest.main()
