import os
import tempfile
import unittest

from app import app


class UploadResponseTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.client.testing = True

    def test_upload_returns_frontend_ready_payload(self):
        csv_content = b"river,station,district,level,danger,trend\nTest River,Station A,Test District,5.2,6.0,Rising\n"
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
            tmp.write(csv_content)
            tmp_path = tmp.name

        try:
            with open(tmp_path, "rb") as fh:
                response = self.client.post(
                    "/upload",
                    data={"file": (fh, "sample.csv")},
                    content_type="multipart/form-data",
                )
        finally:
            os.remove(tmp_path)

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertIn("stats", payload)
        self.assertIn("detected", payload)
        self.assertIn("cumulative", payload)
        self.assertIn("monthly", payload)
        self.assertIn("alerts", payload)
        self.assertIn("district", payload)
        self.assertIn("session", payload)
        self.assertIn("recent", payload)


if __name__ == "__main__":
    unittest.main()
