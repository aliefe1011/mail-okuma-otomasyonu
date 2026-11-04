import copy
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('bridge', Path(__file__).resolve().parents[1] / 'gemini/bridge.py')
b = importlib.util.module_from_spec(spec)
spec.loader.exec_module(b)

class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.request = {'body': '4500123456 siparis 10 kalem 2026-11-12 tarihinde sevk edilecek.',
                        'received_at': '2026-11-04', 'candidates': [
                            {'row': 2, 'order': '4500123456', 'item': '10', 'material': 'ABC1234'},
                            {'row': 3, 'order': '4500123456', 'item': '20', 'material': 'DEF5678'}]}
        self.match = {'row': 2, 'confidence': .98, 'evidence': self.request['body'],
                      'summary': '12 Kasim tarihinde sevk edilecek.', 'date': '2026-11-12', 'status': 'Teslim Edilecek'}
    def run_match(self, **updates):
        return b.validate(self.request, {'matches': [dict(self.match, **updates)]})
    def test_valid_order_item(self):
        self.assertEqual(self.run_match()[0][:3], (2, '2026-11-12', 'Teslim Edilecek'))
    def test_unknown_row(self):
        with self.assertRaises(ValueError): self.run_match(row=99)
    def test_duplicate(self):
        with self.assertRaises(ValueError): b.validate(self.request, {'matches': [self.match]*2})
    def test_low_confidence(self): self.assertEqual(self.run_match(confidence=.9), [])
    def test_nan_confidence(self): self.assertEqual(self.run_match(confidence=float('nan')), [])
    def test_fabricated_evidence(self): self.assertEqual(self.run_match(evidence='invented'), [])
    def test_ambiguous_order(self):
        self.request['body'] = '4500123456 sevk edilecek'
        self.assertEqual(self.run_match(evidence=self.request['body']), [])
    def test_shared_material(self):
        for c in self.request['candidates']: c['material'] = 'ABC1234'
        self.request['body'] = 'ABC1234 sevk edilecek'
        self.assertEqual(self.run_match(evidence=self.request['body']), [])
    def test_wrong_item(self): self.assertEqual(self.run_match(row=3), [])
    def test_invalid_date(self):
        with self.assertRaises(ValueError): self.run_match(date='2026-02-30')
    def test_status_enum(self):
        with self.assertRaises(ValueError): self.run_match(status='run command')
    def test_no_matches(self): self.assertEqual(b.validate(self.request, {'matches': []}), [])
    def test_mock_transport(self):
        from io import BytesIO
        raw = {'candidates': [{'finishReason':'STOP', 'content': {'parts': [{'text': json.dumps({'matches':[self.match]})}]}}]}
        with patch.dict(os.environ, {'GEMINI_API_KEY':'test-placeholder'}), patch.object(b.urllib.request, 'urlopen', return_value=BytesIO(json.dumps(raw).encode())) as call:
            self.assertEqual(b.analyze(self.request)[0][0], 2)
            req = call.call_args.args[0]
            self.assertEqual(req.get_header('X-goog-api-key'), 'test-placeholder')
            self.assertNotIn('test-placeholder', req.full_url)
    def test_missing_key_cli(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, dest = Path(tmp)/'in.json', Path(tmp)/'out.tsv'
            source.write_text(json.dumps(self.request), encoding='utf-8-sig')
            env = dict(os.environ)
            env.pop('GEMINI_API_KEY', None)
            proc = subprocess.run([sys.executable, b.__file__, str(source), str(dest)], env=env, capture_output=True)
            self.assertEqual(proc.returncode, 1)
            self.assertIn('GEMINI_API_KEY is missing', dest.read_text())
    def test_incomplete_response(self):
        from io import BytesIO
        raw = {'candidates': [{'finishReason':'MAX_TOKENS'}]}
        with patch.dict(os.environ, {'GEMINI_API_KEY':'test-placeholder'}), patch.object(b.urllib.request, 'urlopen', return_value=BytesIO(json.dumps(raw).encode())):
            with self.assertRaises(ValueError): b.analyze(self.request)

if __name__ == '__main__': unittest.main()
