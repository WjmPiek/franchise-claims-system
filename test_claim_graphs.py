import unittest
from unittest.mock import patch
from claim_graphs import graphs,SPECS,geometry
from claim_insights import classify

class ClaimGraphsTest(unittest.TestCase):
    dates={'date_from':'2026-08-01','date_to':'2026-09-30'}
    def fixture(self):
        rows=[]
        for product,count in [('A',6),('B',5),('C',4),('D',3),('E',2),('F',1)]:
            for i in range(count):
                rows.append(dict(product=product,relation=['Main Life','Extended Family','Spouse','Children'][i%4],age=['0 - 65','66 - 75'][i%2],gender=['Male','Female','Other','Other / Unknown'][i%4],cause=['Natural','Unnatural','Other'][i%3],reported_on='2026-09-05'))
        return rows
    def test_top_five_ranking_and_every_relation_count_reconcile(self):
        records=self.fixture()
        with patch('claim_graphs.populations',return_value=(records,[],[])):
            graph=graphs(None,self.dates,None,['claim_relationship_products'])[0]
        self.assertEqual(graph['count'],20);self.assertEqual(graph['population_count'],21)
        self.assertEqual([r['label'] for r in graph['rows']],['P1 · A','P2 · B','P3 · C','P4 · D','P5 · E'])
        self.assertTrue(all(sum(r['values'])==r['total'] for r in graph['rows']))
        self.assertEqual(len({s['colour'] for s in graph['series']}),4)
        self.assertEqual(sum(sum(p['value'] for p in s['points']) for s in graph['series']),20)
    def test_gender_and_cause_keep_unknowns_and_do_not_classify_unnatural_as_natural(self):
        records=self.fixture()
        with patch('claim_graphs.populations',return_value=(records,[],[])):
            result=graphs(None,self.dates,{'Allowed'},['claim_gender','claim_cause'])
        for graph in result:
            self.assertEqual(graph['count'],21)
            self.assertEqual(sum(sum(p['value'] for p in s['points']) for s in graph['series']),21)
            self.assertEqual(len({s['colour'] for s in graph['series']}),len(graph['series']))
            self.assertEqual([r['label'] for r in graph['rows']],['2026-08','2026-09'])
        cause=result[1];counts={s['label']:sum(p['value'] for p in s['points']) for s in cause['series']}
        self.assertEqual(counts,{'Natural':9,'Unnatural':7,'Other / Unknown':5})
    def test_product_filter_applies_to_age_series_and_breakdown(self):
        with patch('claim_graphs.populations',return_value=(self.fixture(),[],[])) as population:
            graph=graphs(None,self.dates,{'Branch'},['claim_product_age'],product='B')[0]
        population.assert_called_once_with(None,self.dates,{'Branch'})
        self.assertEqual(graph['count'],5);self.assertEqual(sum(r['count'] for r in graph['product_age_rows']),5)
        self.assertEqual({r['product'] for r in graph['product_age_rows']},{'B'})
        self.assertEqual(sum(sum(p['value'] for p in s['points']) for s in graph['series']),5)
        with patch('claim_graphs.populations',return_value=(self.fixture(),[],[])):
            all_graph=graphs(None,self.dates,None,['claim_product_age'])[0]
        colours={s['label']:s['colour'] for s in all_graph['series']}
        self.assertTrue(all(s['colour']==colours[s['label']] for s in graph['series']))
    def test_empty_or_missing_source_is_explicit(self):
        with patch('claim_graphs.populations',return_value=([],[],['Unanalysed.xlsx'])):
            result=graphs(None,self.dates,None,[s[0] for s in SPECS])
        self.assertTrue(all(not r['chart']['labels'] and r['count']==0 and r['missing'] for r in result))

if __name__=='__main__':unittest.main()
