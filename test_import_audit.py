import csv,io,json,tempfile,unittest
from pathlib import Path
from flask import Flask,g
from direct_data import register_direct_data
from direct_analytics import DirectStore
from import_audit import register_import_audit
from test_direct_data import record

class ImportAuditTest(unittest.TestCase):
    def test_replacement_stream_failure_and_owner_scoped_history(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);app=Flask(__name__,template_folder=str(Path(__file__).parent/'templates'));app.secret_key='test'
            who={'id':1,'role':'admin'}
            @app.before_request
            def user():g.user=who
            app.add_url_rule('/dashboard','dashboard',lambda:'dashboard',methods=['GET','POST'])
            app.add_url_rule('/overview/<section>','overview_page',lambda section:'overview')
            register_direct_data(app,root);audit=register_import_audit(app,root);client=app.test_client()
            def upload(risk,policy='POL001',stream=False):
                output=io.StringIO();csv.writer(output).writerow(record(risk=risk,policy=policy))
                response=client.post('/direct-data',data={'data_month':'2026-09','direct_file':(io.BytesIO(output.getvalue().encode()),'PolicyData_20260901_to_20260930.csv')},headers={'X-Import-Progress':'stream'} if stream else {},buffered=True)
                response.get_data();return response
            self.assertTrue(upload('100').json['ok'])
            streamed=upload('200',stream=True);events=[json.loads(line) for line in streamed.data.splitlines()];self.assertTrue(events[-1]['ok'])
            store=DirectStore(root/'direct_analytics.sqlite');self.assertEqual(store.overview(1)['categories']['source']['rows'],1);self.assertEqual(store.overview(1)['categories']['source']['original_risk'],200)
            failed=upload('300',policy='');self.assertEqual(failed.status_code,400)
            self.assertEqual(store.overview(1)['categories']['source']['original_risk'],200)
            logs=audit.rows(1);self.assertEqual(len(logs),3)
            self.assertEqual(sorted(row['status'] for row in logs),['failed / review','saved / replaced','saved / replaced'])
            self.assertEqual(len(set(row['sha256'] for row in logs)),3)
            self.assertTrue(all(len(row['sha256'])==64 for row in logs))
            self.assertEqual(client.get('/imports/log').status_code,200)
            response=client.get('/imports/log?format=csv');self.assertIn(b'PolicyData',response.data)
            who['id']=2;self.assertEqual(audit.rows(2),[])
            self.assertNotIn(b'PolicyData',client.get('/imports/log?format=csv').data)
            who['role']='viewer';self.assertEqual(client.get('/imports/log').status_code,403)

if __name__=='__main__':unittest.main()
