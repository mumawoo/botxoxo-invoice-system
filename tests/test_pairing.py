import unittest

from invoice_system.models import InvoiceRecord
from invoice_system.pairing import pair_invoice_payment_slips


class PairingTests(unittest.TestCase):
    def test_merges_payment_slip_and_calculates_tip(self):
        invoice = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            expense_amount=100,
        )
        payment = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="Cafe Xuan",
            currency="MXN",
            total_amount=126,
            contents="Pago tarjeta propina",
            crop_image="payment.jpg",
        )

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 1)
        self.assertEqual(paired[0].total_amount, 126)
        self.assertEqual(paired[0].tips, 10)
        self.assertEqual(paired[0].expense_amount, 110)
        self.assertIn("Combined payment slip", paired[0].remarks)
        self.assertEqual(paired[0].supporting_crop_images, ["payment.jpg"])

    def test_merges_payment_slip_even_when_payment_arrives_first(self):
        payment = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="Cafe Xuan",
            currency="MXN",
            total_amount=126,
            contents="CARD PAYMENT TIP",
        )
        invoice = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            expense_amount=100,
        )

        paired = pair_invoice_payment_slips([payment, invoice])

        self.assertEqual(len(paired), 1)
        self.assertEqual(paired[0].seller, "CAFE XUAN")
        self.assertEqual(paired[0].total_amount, 126)
        self.assertEqual(paired[0].tips, 10)

    def test_pairs_payment_slip_with_normalized_date_seller_and_currency(self):
        seller = "Caf" + chr(0xE9) + " Xu" + chr(0xE1) + "n"
        invoice = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="Cafe Xuan",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            expense_amount=100,
        )
        payment = InvoiceRecord(
            invoice_date="12/06/2026",
            seller=seller,
            currency="M.N.",
            total_amount=126,
            contents="Pago con tarjeta propina",
        )

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 1)
        self.assertEqual(paired[0].total_amount, 126)
        self.assertEqual(paired[0].tips, 10)
        self.assertIn("Combined payment slip", paired[0].remarks)

    def test_removes_duplicate_payment_slip_when_amounts_match(self):
        invoice = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116, vat_amount=16)
        payment = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116)

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 1)
        self.assertEqual(paired[0].tips, 0)
        self.assertIn("duplicate payment slip", paired[0].remarks)

    def test_does_not_merge_tip_difference_without_payment_hint(self):
        invoice = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116, vat_amount=16)
        payment = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=126)

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 2)

    def test_does_not_merge_different_merchants(self):
        invoice = InvoiceRecord(invoice_date="2026-06-12", seller="DOANS RESTAURANT", currency="USD", total_amount=27.42, vat_amount=1)
        payment = InvoiceRecord(invoice_date="2026-06-12", seller="DOLLAR GENERAL", currency="USD", total_amount=11, contents="card payment")

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 2)

    def test_removes_duplicate_invoice_photo_when_contents_match(self):
        first = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            contents="Tacos Agua Folio 123",
        )
        duplicate = InvoiceRecord(
            invoice_date="12/06/2026",
            seller="Cafe Xuan",
            currency="M.N.",
            total_amount=116,
            vat_amount=16,
            contents="Tacos Agua Folio 123",
        )

        paired = pair_invoice_payment_slips([first, duplicate])

        self.assertEqual(len(paired), 1)
        self.assertIn("Duplicate invoice photo removed", paired[0].remarks)

    def test_keeps_same_amount_invoice_when_contents_differ(self):
        first = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            contents="Tacos Agua Folio 123",
        )
        second = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            contents="Cafe Pan Folio 999",
        )

        paired = pair_invoice_payment_slips([first, second])

        self.assertEqual(len(paired), 2)

    def test_review_mode_marks_possible_pair_without_merging_or_deleting(self):
        payment = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="Cafe Xuan",
            currency="MXN",
            total_amount=126,
            contents="Pago tarjeta propina",
        )
        invoice = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            expense_amount=100,
        )

        paired = pair_invoice_payment_slips([payment, invoice], mode="review")

        self.assertEqual(len(paired), 2)
        self.assertEqual(paired[0].seller, "CAFE XUAN")
        self.assertEqual(paired[0].total_amount, 116)
        self.assertEqual(paired[0].tips, 0)
        self.assertEqual(paired[1].total_amount, 126)
        self.assertIn("Possible pair PAIR-001", paired[0].remarks)
        self.assertIn("human review required", paired[1].remarks)

    def test_review_mode_does_not_remove_duplicate_content_records(self):
        first = InvoiceRecord(
            invoice_date="2026-06-12",
            seller="CAFE XUAN",
            currency="MXN",
            total_amount=116,
            vat_amount=16,
            contents="Tacos Agua Folio 123",
        )
        duplicate = InvoiceRecord(
            invoice_date="12/06/2026",
            seller="Cafe Xuan",
            currency="M.N.",
            total_amount=116,
            vat_amount=16,
            contents="Tacos Agua Folio 123",
        )

        paired = pair_invoice_payment_slips([first, duplicate], mode="review")

        self.assertEqual(len(paired), 2)

    def test_keeps_unrelated_records(self):
        invoice = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116, vat_amount=16)
        other = InvoiceRecord(invoice_date="2026-06-13", seller="CAFE XUAN", currency="MXN", total_amount=126)

        paired = pair_invoice_payment_slips([invoice, other])

        self.assertEqual(len(paired), 2)


if __name__ == "__main__":
    unittest.main()
