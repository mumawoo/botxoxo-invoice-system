import unittest

from invoice_system.expense_categories import EXPENSE_CATEGORIES, normalize_expense_category


class ExpenseCategoryTests(unittest.TestCase):
    def test_allowed_categories_are_simple_english_list(self):
        self.assertEqual(
            EXPENSE_CATEGORIES,
            (
                "Food",
                "Gas",
                "Car repair",
                "Toll/Parking",
                "Utilities",
                "Internet",
                "Phone",
                "Office supplies",
                "Hotel",
                "Flight",
                "Other",
            ),
        )

    def test_normalizes_food_variants_to_food(self):
        self.assertEqual(normalize_expense_category("restaurant"), "Food")
        self.assertEqual(normalize_expense_category("food_and_beverage"), "Food")
        self.assertEqual(normalize_expense_category("Food & Beverage"), "Food")
        self.assertEqual(normalize_expense_category("", "Codie"), "Food")
        self.assertEqual(normalize_expense_category("", "CORDIA CAFE Y PLANTAS"), "Food")
        self.assertEqual(normalize_expense_category("", "cordial"), "Food")
        self.assertEqual(normalize_expense_category("", "Walmart"), "Food")
        self.assertEqual(normalize_expense_category("", "WAL-MART Supercenter"), "Food")
        self.assertEqual(normalize_expense_category("餐饮"), "Food")

    def test_normalizes_common_business_categories(self):
        self.assertEqual(normalize_expense_category("", "Pemex gasolina"), "Gas")
        self.assertEqual(normalize_expense_category("", "Taller mecanico llantas"), "Car repair")
        self.assertEqual(normalize_expense_category("", "Caseta peaje autopista"), "Toll/Parking")
        self.assertEqual(normalize_expense_category("", "CFE electricidad"), "Utilities")
        self.assertEqual(normalize_expense_category("", "Telmex internet fibra"), "Internet")
        self.assertEqual(normalize_expense_category("", "Telcel telefono celular"), "Phone")
        self.assertEqual(normalize_expense_category("", "Office Depot papel"), "Office supplies")
        self.assertEqual(normalize_expense_category("", "Hotel City Express"), "Hotel")
        self.assertEqual(normalize_expense_category("", "Volaris vuelo"), "Flight")

    def test_unknown_defaults_to_other(self):
        self.assertEqual(normalize_expense_category("", "unknown merchant"), "Other")


if __name__ == "__main__":
    unittest.main()
