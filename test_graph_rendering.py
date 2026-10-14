import unittest
from central_graphs import chart_geometry

class GraphRenderingTest(unittest.TestCase):
    def test_single_month_is_centred(self):
        c=chart_geometry([{'label':'2026-09','value':100}])
        self.assertEqual(c['points'][0]['x'],300)
        self.assertLess(c['points'][0]['y'],c['baseline'])
    def test_zero_and_negative_values_have_valid_geometry(self):
        for values in ([0],[0,0],[-20,30]):
            c=chart_geometry([{'label':str(i),'value':v} for i,v in enumerate(values)])
            self.assertTrue(all(20<=p['y']<=190 for p in c['points']))
            self.assertNotIn('nan',c['path'])
    def test_empty_period_is_explicit(self):
        self.assertEqual(chart_geometry([])['points'],[])

if __name__=='__main__':unittest.main()
