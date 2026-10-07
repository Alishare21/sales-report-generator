import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
from audit import save_audit, finalize_review


class AuditTests(unittest.TestCase):
    def test_append_only_and_fresh_resolution(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);output=root/'outputs'/'sample';output.mkdir(parents=True)
            verify=output/'verification.json'
            verify.write_text(json.dumps({'checks':{'Visual review':{'done':False,'evidence':'pending'}}}))
            first=save_audit(output,audit_root=root/'audits')
            first_bytes=first.read_bytes()
            verify.write_text(json.dumps({'checks':{'Visual review':{'done':True,'evidence':'inspected'}}}))
            second=save_audit(output,'reviewed',audit_root=root/'audits')
            self.assertNotEqual(first,second);self.assertEqual(first.read_bytes(),first_bytes)
            record=json.loads(second.with_suffix('.json').read_text())
            self.assertEqual(record['transitions'],[{'id':'CHECK-01','status':'resolved'}])
    def test_failed_build_preserves_gap_without_claiming_resolution(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);output=root/'sample';output.mkdir()
            (output/'verification.json').write_text(json.dumps({'checks':{'Visual':{'done':False}}}))
            save_audit(output,audit_root=root/'audits')
            failed=save_audit(output,'failed',error='load mismatch',audit_root=root/'audits')
            record=json.loads(failed.with_suffix('.json').read_text())
            self.assertIn({'id':'CHECK-01','status':'not rechecked'},record['transitions'])
    def test_stale_visual_review_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);(root/'findings_report.pdf').write_bytes(b'changed')
            review=root/'review.json';review.write_text(json.dumps({'artifact_sha256':{'pdf':'old'}}))
            with self.assertRaises(ValueError):finalize_review(root,review)


if __name__=='__main__':unittest.main()
