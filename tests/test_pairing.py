import unittest

from invoice_system.models import InvoiceRecord
from invoice_system.pairing import is_possible_duplicate, mark_possible_matches_with_protected, pair_invoice_payment_slips


class PairingTests(unittest.TestCase):
    def test_protected_match_is_warned_but_not_combined(self):
        protected = InvoiceRecord(
            invoice_date="2026-05-16",
            seller="Las Palmas",
            currency="MXN",
            total_amount=536,
            vat_amount=73.93,
            crop_image="037_receipt.jpg",
        )
        payment = InvoiceRecord(
            invoice_date="2026-05-16",
            seller="REST HOTEL LAS PALMAS",
            currency="MXN",
            total_amount=616.40,
            contents="VENTA CON PROPINA",
            crop_image="163_payment.jpg",
        )

        result = mark_possible_matches_with_protected([payment], [protected])

        self.assertEqual(result, [payment])
        self.assertEqual(payment.total_amount, 616.40)
        self.assertIn("Possible pair with protected crop 037", payment.remarks)
        self.assertIn("payment difference 80.40", payment.remarks)

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

    def test_merges_no_vat_restaurant_receipt_with_bank_tip_slip(self):
        invoice = InvoiceRecord(
            invoice_date="2026-05-09",
            seller="SUSHI ROLL S.A. DE C.V.",
            currency="MXN",
            total_amount=566,
            expense_amount=566,
            contents="1 PEPSI LIGHT, 1 TEPPANYAKI MIX ESP, 1 MATCHA CAKE",
            remarks="Mesa # 65, 1 Personas, Atendio: Roman Bautista",
            crop_image="030.jpg",
        )
        payment = InvoiceRecord(
            invoice_date="2026-05-09",
            seller="MIFEL SUSHI ROLL CUMBRES",
            currency="MXN",
            total_amount=622.6,
            expense_amount=566,
            tips=56.6,
            contents="SUSHI ROLL CUMBRES",
            remarks="VENTA CON PROPINA, Monterrey NL, CREDITO/HSBC MEX/MASTERCARD",
            crop_image="031.jpg",
        )

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 1)
        self.assertEqual(paired[0].seller, "SUSHI ROLL S.A. DE C.V.")
        self.assertEqual(paired[0].total_amount, 622.6)
        self.assertEqual(paired[0].tips, 56.6)
        self.assertEqual(paired[0].supporting_crop_images, ["031.jpg"])

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

    def test_does_not_merge_same_amount_without_payment_hint(self):
        invoice = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116, vat_amount=16)
        payment = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116)

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 2)

    def test_merges_same_amount_when_payment_hint_is_present(self):
        invoice = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116, vat_amount=16)
        payment = InvoiceRecord(invoice_date="2026-06-12", seller="CAFE XUAN", currency="MXN", total_amount=116, contents="card payment")

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

    def test_same_photo_exact_reported_tip_can_pair_despite_bank_merchant_alias(self):
        source = "telegram_photo.jpg"
        invoice = InvoiceRecord(
            invoice_date="2026-06-16", seller="supersalads", currency="MXN", total_amount=233,
            vat_amount=32.14, contents="Louisiana chicken bowl", source_image=source, crop_image="200_receipt.jpg",
        )
        payment = InvoiceRecord(
            invoice_date="2026-06-16", seller="REST SS CITAOTINA ESC", currency="MXN", total_amount=256.30,
            tips=23.30, contents="VENTA CON PROPINA", source_image=source, crop_image="201_card.jpg",
        )

        paired = pair_invoice_payment_slips([invoice, payment])

        self.assertEqual(len(paired), 1)
        self.assertEqual(paired[0].total_amount, 256.30)
        self.assertEqual(paired[0].tips, 23.30)
        self.assertEqual(paired[0].supporting_crop_images, ["201_card.jpg"])

    def test_exact_tip_does_not_override_different_source_and_merchant(self):
        invoice = InvoiceRecord(
            invoice_date="2026-06-16", seller="supersalads", currency="MXN", total_amount=233,
            vat_amount=32.14, source_image="one.jpg",
        )
        payment = InvoiceRecord(
            invoice_date="2026-06-16", seller="DIFFERENT STORE", currency="MXN", total_amount=256.30,
            tips=23.30, contents="VENTA CON PROPINA", source_image="two.jpg",
        )

        self.assertEqual(len(pair_invoice_payment_slips([invoice, payment])), 2)

    def test_marks_only_new_same_content_record_as_possible_duplicate(self):
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

        self.assertEqual(len(paired), 2)
        self.assertNotIn("Possible duplicate", paired[0].remarks)
        self.assertIn("Possible duplicate", paired[1].remarks)

    def test_marks_new_same_amount_invoice_when_contents_differ(self):
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
        self.assertNotIn("Possible duplicate", paired[0].remarks)
        self.assertIn("Possible duplicate", paired[1].remarks)

    def test_missing_date_still_warns_for_same_merchant_currency_and_amount(self):
        dated = InvoiceRecord(
            invoice_date="2026-06-06",
            seller="Nota De Cuenta",
            currency="MXN",
            total_amount=150,
            crop_image="111_old.jpg",
        )
        undated = InvoiceRecord(
            invoice_date="",
            seller="Nota De Cuenta",
            currency="MXN",
            total_amount=150,
            crop_image="207_new.jpg",
        )

        self.assertTrue(is_possible_duplicate(dated, undated))

    def test_different_known_dates_are_not_possible_duplicates(self):
        first = InvoiceRecord(
            invoice_date="2026-06-06",
            seller="Nota De Cuenta",
            currency="MXN",
            total_amount=150,
            crop_image="111_old.jpg",
        )
        second = InvoiceRecord(
            invoice_date="2026-06-07",
            seller="Nota De Cuenta",
            currency="MXN",
            total_amount=150,
            crop_image="207_new.jpg",
        )

        self.assertFalse(is_possible_duplicate(first, second))

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
