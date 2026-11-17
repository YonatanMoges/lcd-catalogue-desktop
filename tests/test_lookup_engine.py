"""Regression tests for the public, fictional LCD catalogue demo.

Run from the repository root:
    python -m unittest discover -s tests -v
"""
import json
import unittest
from pathlib import Path

from lookup_engine import lookup, norm, number, word_order_match


DEMO = Path(__file__).resolve().parents[1] / "examples" / "demo_catalogue.json"


class CatalogueLookupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalogue = json.loads(DEMO.read_text(encoding="utf-8"))

    def test_normalization_ignores_spaces_and_hyphens(self):
        self.assertEqual(norm(" DPA-100 "), "DPA100")

    def test_reordered_names_require_the_same_words(self):
        self.assertTrue(word_order_match("DemoPhone A1", "A1 DemoPhone"))
        self.assertFalse(word_order_match("DemoPhone A1", "DemoPhone A2"))

    def test_rejects_missing_and_invalid_numbers(self):
        for value in (None, "", "not a number", float("nan")):
            self.assertIsNone(number(value))

    def test_exact_phone_search_finds_local_stock(self):
        result = lookup(self.catalogue, "DemoPhone A1")
        self.assertEqual(result["status"], "MATCH_FOUND")
        self.assertEqual(result["match_count"], 1)
        self.assertEqual(result["products"][0]["store_quantity"], 12)
        self.assertEqual(result["products"][0]["main_selling_price"], 24.5)
        self.assertTrue(result["automatic_quote_allowed"])

    def test_compatible_model_code_search(self):
        result = lookup(self.catalogue, "DPA100")
        self.assertEqual(result["match_count"], 1)
        self.assertIn("DPA100", result["products"][0]["product"])

    def test_out_of_stock_product_requires_warehouse_confirmation(self):
        result = lookup(self.catalogue, "DemoPhone B2")
        self.assertEqual(result["status"], "WAREHOUSE_OR_SHORTAGE_REVIEW")
        self.assertEqual(result["products"][0]["source"], "BIG_WAREHOUSE_INDICATED")
        self.assertFalse(result["automatic_quote_allowed"])

    def test_insufficient_local_stock_requires_review(self):
        result = lookup(self.catalogue, "DemoPhone A1", quantity=13)
        self.assertIn("INSUFFICIENT_LOCAL_STOCK", result["products"][0]["review_reasons"])
        self.assertFalse(result["automatic_quote_allowed"])

    def test_missing_selling_price_blocks_quotation(self):
        result = lookup(self.catalogue, "DemoPhone C3")
        product = result["products"][0]
        self.assertEqual(result["status"], "SELLING_PRICE_REVIEW")
        self.assertIsNone(product["main_selling_price"])
        self.assertFalse(product["customer_quote_allowed"])
        self.assertFalse(result["automatic_quote_allowed"])

    def test_unknown_phone_returns_no_products(self):
        result = lookup(self.catalogue, "UnknownPhone Z99")
        self.assertEqual(result["status"], "NOT_FOUND_IN_LOADED_SHEETS")
        self.assertEqual(result["products"], [])

    def test_blank_query_is_rejected(self):
        with self.assertRaises(ValueError):
            lookup(self.catalogue, "   ")

    def test_nonpositive_quantity_requires_review(self):
        result = lookup(self.catalogue, "DemoPhone A1", quantity=0)
        self.assertIsNone(result["requested_quantity"])
        self.assertTrue(result["needs_review"])


if __name__ == "__main__":
    unittest.main()
