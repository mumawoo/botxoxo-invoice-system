import unittest
from pathlib import Path

from invoice_system.config import Settings
from invoice_system.visual_count import AIVisualCounter


class VisualCountTests(unittest.TestCase):
    def test_openai_visual_count_is_removed(self):
        result = AIVisualCounter(Settings(openai_api_key="key", ai_visual_count_enabled=True)).count(Path("photo.jpg"))

        self.assertIsNone(result.count)
        self.assertIn("removed", result.reason)


if __name__ == "__main__":
    unittest.main()
