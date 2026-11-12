from __future__ import annotations

import unittest
from pathlib import Path

from proposal_engine.booklet_plan_design import booklet_plan_design_extractor as extractor
from proposal_engine.booklet_plan_design.labeled_dental_rate_mapper import (
    CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION,
    canonical_dental_provision_code,
    map_labeled_dental_coverage_statuses,
    map_labeled_dental_maximums,
    map_labeled_dental_rates,
)


class LabeledDentalRateMapperTests(unittest.TestCase):
    def test_dotted_leader_layout_maps_each_named_rate(self) -> None:
        text = """
        [[PAGE 18]]
        Dental Care Coverage
        % Payment of Allowable Expenses
        Basic Dental Care ........................................................ 100%
        Examinations and Diagnostic Services
        Preventive Services
        Routine Dental Care (subject to certain limits) .......................... 100%
        Minor Restorative Services, Endodontics, Periodontics and Oral Surgery
        Dental Restorative Services (subject to certain limits) ................... 60%
        Major Restorative Services and Fixed Prosthodontics
        Orthodontic Care (subject to certain limits) .............................. 60%
        Orthodontics - Adults & Dependent Children
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}

        self.assertEqual("100%", rates["basic"].reimbursement_text)
        self.assertEqual("100%", rates["routine"].reimbursement_text)
        self.assertEqual("60%", rates["major"].reimbursement_text)
        self.assertEqual("60%", rates["orthodontic"].reimbursement_text)
        self.assertNotIn("preventive", rates)
        self.assertEqual(18, rates["major"].source_page)

    def test_level_layout_maps_basic_dentures_and_major_with_combined_maximums(self) -> None:
        text = """
        [[PAGE 22]]
        Dental
        Level I - Basic Services
        Includes items such as complete oral exam and x-rays.
        80% to a combined maximum of $2,500 per calendar year
        [[PAGE 23]]
        Level III - Dentures
        Includes initial provision and replacement of removable dentures.
        50% to a combined maximum of $2,500 per calendar year
        Level IV - Major Restorative Services
        Includes crowns, onlays, inlays and fixed bridgework.
        50% to a combined maximum of $2,500 per calendar year
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}

        self.assertEqual("80%", rates["basic"].reimbursement_text)
        self.assertEqual("50%", rates["dentures"].reimbursement_text)
        self.assertEqual("50%", rates["major"].reimbursement_text)
        self.assertEqual("To a combined maximum of $2,500 per calendar year", rates["major"].maximum)
        self.assertEqual(23, rates["dentures"].source_page)

    def test_reimbursement_table_maps_four_rates_and_keeps_ortho_eligibility(self) -> None:
        text = """
        [[PAGE 31]]
        Dental Care - Contract Number 152628
        Reimbursement Level
        Preventive procedures 80%
        Basic procedures 80%
        Major procedures 50%
        Orthodontic procedures 50%, only for children under age 20
        Maximum benefit
        Benefit year maximum $2,000 per person
        Lifetime maximum Orthodontic procedures - $3,000 per person
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}

        self.assertEqual("80%", rates["preventive"].reimbursement_text)
        self.assertEqual("80%", rates["basic"].reimbursement_text)
        self.assertEqual("50%", rates["major"].reimbursement_text)
        self.assertEqual("50%", rates["orthodontic"].reimbursement_text)
        self.assertEqual("only for children under age 20", rates["orthodontic"].eligibility.lower())
        self.assertEqual("", rates["orthodontic"].maximum)

    def test_explicit_type_legend_maps_separated_rates_and_maximums(self) -> None:
        text = """
        [[PAGE 42]]
        Dental
        Type A - Basic services.
        Recall examination period: once in any period of 9 months.
        This Dental plan includes the following Basic services options.
        Type B - Major restorative services:
        This Dental plan includes the following Major restorative options.
        Type C - Orthodontic services:
        Only dependent children under age 18 are eligible for Orthodontic services.
        Reimbursement percentage:
        Type A: 80%
        Type B: 50%
        Type C: 50%
        Maximum amount:
        Annual calendar year maximum for Type A and Type B combined: $2,500
        Lifetime maximum for Type C: $3,000
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}

        self.assertEqual("80%", rates["basic"].reimbursement_text)
        self.assertEqual("50%", rates["major"].reimbursement_text)
        self.assertEqual("50%", rates["orthodontic"].reimbursement_text)
        self.assertEqual(
            "Annual calendar year maximum for Type A and Type B combined: $2,500",
            rates["basic"].maximum,
        )
        self.assertEqual(rates["basic"].maximum, rates["major"].maximum)
        self.assertEqual("Lifetime maximum for Type C: $3,000", rates["orthodontic"].maximum)
        self.assertEqual(
            "dependent children under age 18",
            rates["orthodontic"].eligibility.lower(),
        )

    def test_coverage_layout_maps_accidental_injury_and_exact_labeled_maximums(self) -> None:
        text = """
        [[PAGE 51]]
        Dentalcare
        Reimbursement Levels
        Basic Coverage                         100%
        Major Coverage                          60%
        Orthodontic Coverage                     60%
        Accidental Dental Injury Coverage       100%
        Plan Maximums
        Accidental Dental Injury Treatment      Unlimited
        Orthodontic Treatment                   $2,750 lifetime
        All Other Treatment                     Unlimited
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}
        maximums = {item.provision_code: item for item in map_labeled_dental_maximums(text)}

        self.assertEqual("100%", rates["basic"].reimbursement_text)
        self.assertEqual("60%", rates["major"].reimbursement_text)
        self.assertEqual("60%", rates["orthodontic"].reimbursement_text)
        self.assertEqual("100%", rates["accidental_injury"].reimbursement_text)
        self.assertEqual("$2,750 lifetime", rates["orthodontic"].maximum)
        self.assertEqual("Unlimited", rates["accidental_injury"].maximum)
        self.assertEqual("Unlimited", maximums["other_treatment"].maximum)

    def test_empire_schedule_maps_restorative_aliases_and_shared_benefit_period_maximum(self) -> None:
        text = """
        [[PAGE 7]]
        SCHEDULE OF BENEFITS
        DENTAL
        Coinsurance:
        Basic Restorative,
        Periodontics/Endodontics                 80%
        Major Restorative                       50%
        Benefit Period Maximum:
        Basic Restorative,
        Periodontics/Endodontics, Major
        Restorative                             $1,500
        Dental Fee Guide: Current Fee Guide
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}
        maximums = {item.provision_code: item for item in map_labeled_dental_maximums(text)}

        self.assertEqual("80%", rates["basic"].reimbursement_text)
        self.assertEqual("50%", rates["major"].reimbursement_text)
        self.assertEqual("Combined $1,500 per benefit period", rates["basic"].maximum)
        self.assertEqual(rates["basic"].maximum, rates["major"].maximum)
        self.assertEqual(rates["basic"].maximum, maximums["basic"].maximum)
        self.assertNotIn("endodontic_periodontic", maximums)

    def test_rbc_summary_keeps_endodontic_periodontic_separate_and_combines_basic_major_maximum(self) -> None:
        text = """
        [[PAGE 9]]
        GROUP DENTAL CARE (DTL) BENEFIT SUMMARY
        DTL Reimbursement percentage/Benefit Description:
        Basic Dental Services: 100% after deductible.
        Endodontic/Periodontic Services: 100% after deductible.
        Major Dental Services: 50% after deductible.
        DTL Benefit Period Maximum:
        Basic and Major Dental Services (combined): $2,500 per insured.
        DTL Recall Period: Every 9 consecutive months.
        """

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}

        self.assertEqual("100%", rates["basic"].reimbursement_text)
        self.assertEqual("100%", rates["endodontic_periodontic"].reimbursement_text)
        self.assertEqual("50%", rates["major"].reimbursement_text)
        self.assertEqual("Combined $2,500 per insured per benefit period", rates["basic"].maximum)
        self.assertEqual(rates["basic"].maximum, rates["major"].maximum)

    def test_unlabeled_multi_rate_text_is_not_guessed(self) -> None:
        text = "[[PAGE 5]] Dental coverage reimbursement is 80%; 50%."

        self.assertEqual([], map_labeled_dental_rates(text))

    def test_conflicting_values_for_one_label_are_not_canonicalized(self) -> None:
        text = """
        [[PAGE 5]]
        Dental
        Basic Services 80%
        [[PAGE 6]]
        Dental
        Basic Services 90%
        """

        self.assertEqual([], map_labeled_dental_rates(text))

    def test_distant_body_percentage_does_not_override_nearby_schedule_rate(self) -> None:
        text = """
        [[PAGE 12]]
        Schedule of Benefits
        Dental Care Coverage
        Dental Restorative Services ............................................ 50%
        [[PAGE 80]]
        Dental Restorative Services
        """ + ("Detailed policy wording without a rate. " * 40) + "100%"

        rates = {item.provision_code: item for item in map_labeled_dental_rates(text)}

        self.assertEqual("50%", rates["major"].reimbursement_text)

    def test_canonical_rows_ignore_treatment_cost_percentage_and_keep_dental_frequencies_separate(self) -> None:
        text = """
        [[PAGE 24]]
        Dental Care Coverage
        Basic Dental Care ................................................... 100%
        Routine Dental Care ................................................. 100%
        Dental Restorative Services .......................................... 60%
        Orthodontic Care ...................................................... 60%
        Plan Maximums
        Orthodontic Treatment $2,750 lifetime
        [[PAGE 88]]
        Orthodontics
        A maximum amount of 25% of the total treatment cost may be paid when treatment begins.
        """

        rows = extractor.merge_labeled_dental_rate_rows(
            [],
            text,
            Path("grouphealth-booklet.pdf"),
            extractor.ClientInfo("CLIENT", "Example Client", "GroupHEALTH"),
            "All Employees",
        )
        by_benefit = {row.benefit: row for row in rows}

        self.assertNotIn("Dental Routine Care", by_benefit)
        self.assertEqual("100%", by_benefit["Dental Basic Services"].coinsurance)
        self.assertEqual("60%", by_benefit["Dental Major Restorative Services"].coinsurance)
        self.assertEqual("", by_benefit["Dental Major Restorative Services"].frequency)
        self.assertEqual("60%", by_benefit["Dental Orthodontic Services"].coinsurance)
        self.assertEqual("$2,750 lifetime", by_benefit["Dental Orthodontic Services"].maximum)
        self.assertEqual("per lifetime", by_benefit["Dental Orthodontic Services"].frequency)

    def test_shared_maximum_and_not_included_status_are_mapped_by_provision(self) -> None:
        text = """
        [[PAGE 27]]
        Dental Care Coverage
        Basic Dental Care 80%
        Routine Dental Care 80%
        Dental Restorative Services 50%
        Orthodontic Care Not Included
        Maximum Amount Basic Dental Care/Routine Dental Care/Major Restorative Services
        Covered combined maximum of $3,000 per Insured Person per calendar year
        """

        maximums = {item.provision_code: item for item in map_labeled_dental_maximums(text)}
        statuses = {
            item.provision_code: item
            for item in map_labeled_dental_coverage_statuses(text)
        }

        self.assertEqual(
            {"basic", "routine", "major"},
            set(maximums),
        )
        self.assertTrue(
            all(
                item.maximum == "Combined maximum of $3,000 per Insured Person per calendar year"
                for item in maximums.values()
            )
        )
        self.assertEqual("Not Included", statuses["orthodontic"].coverage_status)

        rows = extractor.extract_plan_design_rows(
            text,
            Path("not-included-booklet.pdf"),
            extractor.ClientInfo("CLIENT", "Example Client", "Example Carrier"),
            extractor.load_extraction_rules(),
            "All Employees",
        )
        orthodontic = next(
            row
            for row in rows
            if row.benefit == "Dental Orthodontic Services"
        )
        self.assertEqual("Coverage: Not Included", orthodontic.coverage_description)
        self.assertEqual("", orthodontic.coinsurance)

    def test_covered_first_shared_no_maximum_maps_each_dental_provision(self) -> None:
        text = """
        [[PAGE 18]]
        Paramedical Maximum $500 per practitioner per calendar year
        [[PAGE 24]]
        Dental Care Coverage
        Basic Dental Care ................................................... 100%
        Routine Dental Care ................................................. 100%
        Dental Restorative Services .......................................... 60%
        Orthodontic Care ...................................................... 60%
        Maximum Amount Covered
        Basic Dental Care/Routine Dental Care/Major Restorative Services
        No maximum
        Orthodontic Care
        $2,750 per Insured per lifetime
        """

        maximums = {item.provision_code: item for item in map_labeled_dental_maximums(text)}
        rows = extractor.merge_labeled_dental_rate_rows(
            [],
            text,
            Path("grouphealth-booklet.pdf"),
            extractor.ClientInfo("CLIENT", "Example Client", "GroupHEALTH"),
            "All Employees",
        )
        by_benefit = {row.benefit: row for row in rows}

        self.assertEqual("No maximum", maximums["basic"].maximum)
        self.assertEqual("No maximum", maximums["routine"].maximum)
        self.assertEqual("No maximum", maximums["major"].maximum)
        self.assertEqual("No maximum", by_benefit["Dental Basic Services"].maximum)
        self.assertEqual("No maximum", by_benefit["Dental Major Restorative Services"].maximum)
        self.assertEqual("$2,750", by_benefit["Dental Orthodontic Services"].maximum)
        self.assertEqual("per lifetime", by_benefit["Dental Orthodontic Services"].frequency)

    def test_dental_merge_does_not_borrow_a_maximum_from_another_page(self) -> None:
        text = """
        [[PAGE 18]]
        Paramedical Maximum $500 per practitioner per calendar year
        [[PAGE 24]]
        Dental Care Coverage
        Dental Restorative Services .......................................... 60%
        """
        contaminated = extractor.ExtractionRow(
            client_id="CLIENT",
            client_name="Example Client",
            carrier="GroupHEALTH",
            class_name="All Employees",
            benefit="Dental Major Restorative Services",
            coverage_description="Dental Restorative Services 60%",
            coinsurance="60%",
            maximum="$500",
            frequency="",
            waiting_period="",
            termination_age="",
            notes="Generic extraction",
            source_file="grouphealth-booklet.pdf",
            source_section="Dental Care",
            confidence="Medium",
            review_flag="Review",
            last_updated="2026-07-21",
            source_page=24,
        )

        rows = extractor.merge_labeled_dental_rate_rows(
            [contaminated],
            text,
            Path("grouphealth-booklet.pdf"),
            extractor.ClientInfo("CLIENT", "Example Client", "GroupHEALTH"),
            "All Employees",
        )
        major = next(row for row in rows if row.benefit == "Dental Major Restorative Services")

        self.assertEqual("", major.maximum)

    def test_alias_mapping_is_single_provision_only(self) -> None:
        self.assertEqual("basic", canonical_dental_provision_code("Dental Basic Services"))
        self.assertEqual("major", canonical_dental_provision_code("Level IV - Major Restorative Services"))
        self.assertEqual("orthodontic", canonical_dental_provision_code("Orthodontics"))
        self.assertEqual("", canonical_dental_provision_code("Preventive, Basic and Major Restorative Services"))
        self.assertEqual("CANONICAL_LABELED_DENTAL_RATE_MAPPER_V1", CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION)

    def test_booklet_extractor_emits_separate_canonical_rows_for_type_layout(self) -> None:
        text = """
        [[PAGE 42]]
        Dental
        Type A - Basic services.
        Type B - Major restorative services:
        Type C - Orthodontic services:
        Only dependent children under age 18 are eligible for Orthodontic services.
        Reimbursement percentage:
        Type A: 80%
        Type B: 50%
        Type C: 50%
        Maximum amount:
        Annual calendar year maximum for Type A and Type B combined: $2,500
        Lifetime maximum for Type C: $3,000
        """
        client = extractor.ClientInfo("CLIENT", "Example Client", "Example Carrier")

        rows = extractor.extract_plan_design_rows(
            text,
            Path("example-booklet.pdf"),
            client,
            extractor.load_extraction_rules(),
            "All Employees",
        )
        dental_rates = {row.benefit: row for row in rows if row.coinsurance}

        self.assertEqual("80%", dental_rates["Dental Basic Services"].coinsurance)
        self.assertEqual("50%", dental_rates["Dental Major Restorative Services"].coinsurance)
        self.assertEqual("50%", dental_rates["Dental Orthodontic Services"].coinsurance)
        self.assertTrue(
            all("80%; 50%" not in row.coinsurance for row in rows),
            "Labeled values must not be flattened back into one ambiguous percentage field.",
        )

    def test_workbook_reprocess_replaces_stale_combined_dental_rows_idempotently(self) -> None:
        text = """
        [[PAGE 12]]
        Dental
        Basic Services 80%
        Major Services 50%
        """
        client = extractor.ClientInfo("CLIENT", "Example Client", "Example Carrier")
        incoming = extractor.extract_plan_design_rows(
            text,
            Path("current-booklet.pdf"),
            client,
            extractor.load_extraction_rules(),
            "All Employees",
        )
        workbook = extractor.Workbook()
        sheet = workbook.active
        sheet.title = extractor.PLAN_SHEET
        sheet.append(extractor.PLAN_COLUMNS)
        stale = extractor.ExtractionRow(
            client_id="CLIENT",
            client_name="Example Client",
            carrier="Example Carrier",
            class_name="All Employees",
            benefit="Dental Care",
            coverage_description="Dental reimbursement 80%; 50%",
            coinsurance="80%; 50%",
            maximum="",
            frequency="",
            waiting_period="",
            termination_age="",
            notes="",
            source_file="current-booklet.pdf",
            source_section="Dental Care",
            confidence="Medium",
            review_flag="Review",
            last_updated="2026-01-01",
            source_page=12,
        )
        stale_values = extractor.plan_row_to_dict(stale)
        sheet.append([stale_values.get(column, "") for column in extractor.PLAN_COLUMNS])

        first_added, _first_updated, _first_counts = extractor.update_plan_design_sheet(sheet, incoming)
        second_added, _second_updated, _second_counts = extractor.update_plan_design_sheet(sheet, incoming)
        header_map = extractor.worksheet_header_map(sheet)
        saved = [
            (
                sheet.cell(row=row, column=header_map["Benefit"]).value,
                sheet.cell(row=row, column=header_map["Coinsurance"]).value,
            )
            for row in range(2, sheet.max_row + 1)
        ]

        self.assertEqual(2, first_added)
        self.assertEqual(0, second_added)
        self.assertEqual(
            [("Dental Basic Services", "80%"), ("Dental Major Restorative Services", "50%")],
            saved,
        )


if __name__ == "__main__":
    unittest.main()
