import unittest
from decimal import Decimal
from business_overview import branch_totals

class BusinessOverviewTest(unittest.TestCase):
    def test_months_roll_up_to_one_branch_without_recalculating(self):
        rows=[{'franchise':'A','summary':{'source_retail':'200.25','original_risk':'150','payover':'140','rows':2}},
              {'franchise':'A','summary':{'source_retail':'100','original_risk':'80','payover':'70','rows':3}},
              {'franchise':'B','summary':{'source_retail':'50','original_risk':'30','payover':'25','rows':1}}]
        result=branch_totals(rows,[{'franchise':'A','claims':'12.50'},{'franchise':'OUTSIDE SCOPE','claims':'999'}])
        self.assertEqual(len(result),2)
        self.assertEqual(result[0]['retail'],Decimal('300.25'))
        self.assertEqual(result[0]['payover'],Decimal('210'))
        self.assertEqual(result[0]['rows'],5)
        self.assertEqual(result[0]['claims'],Decimal('12.50'))

if __name__=='__main__':unittest.main()
