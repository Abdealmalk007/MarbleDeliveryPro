"""Public-release privacy and authentication regression tests."""
import unittest

from test_v1005 import app_module


class PublicReleaseRegression(unittest.TestCase):
    def test_clean_company_profile(self):
        self.assertEqual(app_module.APP_VERSION, "10.0.5")
        self.assertEqual(app_module.AppSettings().company_name, "")

    def test_empty_password_is_never_valid(self):
        self.assertFalse(app_module.password_matches("", "", ""))
        self.assertFalse(app_module.password_matches("", "a0" * 16, "0" * 64))
        self.assertFalse(app_module.DatabaseManager.authenticate("admin", ""))

    def test_valid_password_authentication(self):
        salt, digest = app_module.password_hash("Strong-Test-Passphrase-01")
        self.assertTrue(app_module.password_matches("Strong-Test-Passphrase-01", salt, digest))
        self.assertFalse(app_module.password_matches("wrong", salt, digest))

    def test_fresh_database_has_no_records_or_private_content(self):
        with app_module.get_db() as conn:
            for table in (
                "records", "attachments", "audit_log", "report_archive",
                "integration_outbox", "deleted_records", "directories",
                "report_templates", "session_state", "settings",
            ):
                with self.subTest(table=table):
                    self.assertEqual(conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0], 0)
            accounts = conn.execute(
                "SELECT username, salt, password_hash FROM users"
            ).fetchall()
            self.assertEqual(accounts, [("admin", "", "")])


if __name__ == "__main__":
    unittest.main()
