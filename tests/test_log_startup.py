import logging
import unittest
from pathlib import Path

import app_runtime


class StartupLogTests(unittest.TestCase):
    def test_clear_startup_log_removes_previous_active_and_rotated_logs(self):
        logger = logging.getLogger("kks-audit")
        existing_handlers = list(logger.handlers)
        for handler in existing_handlers:
            logger.removeHandler(handler)
            handler.close()

        original_log_dir = app_runtime.LOG_DIR
        original_log_path = app_runtime.LOG_PATH
        test_dir = Path(__file__).resolve().parent / ".log-startup-test"
        try:
            test_dir.mkdir(parents=True, exist_ok=True)
            app_runtime.LOG_DIR = test_dir
            app_runtime.LOG_PATH = test_dir / "kks-audit.log"
            app_runtime.LOG_PATH.write_text("old startup\n", encoding="utf-8")
            rotated = test_dir / "kks-audit.log.1"
            rotated.write_text("old rotated\n", encoding="utf-8")

            configured = app_runtime.configure_logging(clear=True)
            for handler in configured.handlers:
                handler.flush()

            self.assertEqual(app_runtime.LOG_PATH.read_text(encoding="utf-8"), "")
            self.assertFalse(rotated.exists())

            configured.info("new startup")
            for handler in configured.handlers:
                handler.flush()
            self.assertIn("new startup", app_runtime.LOG_PATH.read_text(encoding="utf-8"))
        finally:
            for handler in list(logger.handlers):
                logger.removeHandler(handler)
                handler.close()
            if test_dir.is_dir():
                for test_path in test_dir.glob("kks-audit.log*"):
                    test_path.unlink(missing_ok=True)
                test_dir.rmdir()
            for handler in existing_handlers:
                logger.addHandler(handler)
            app_runtime.LOG_DIR = original_log_dir
            app_runtime.LOG_PATH = original_log_path


if __name__ == "__main__":
    unittest.main()
