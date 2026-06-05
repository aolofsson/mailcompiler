"""Unit tests for the contact extraction/query helpers."""

from mailcompiler.mailcompiler import (
    matches, parse_args, build_criteria,
)

CONTACTS = [
    {"first_name": "Dana", "last_name": "Brooks", "company": "Northwind",
     "primary_email": "dana@northwind.com", "emails": ["dana@northwind.com"],
     "num_emails": 5, "num_sent": 1, "num_received": 4,
     "first_interaction": "2022-06-03", "last_interaction": "2022-06-20"},
    {"first_name": "Jordan", "last_name": "Vale", "company": "Globex",
     "primary_email": "jordan@globex.com",
     "emails": ["jordan@globex.com", "jordan.vale@globex.com"],
     "num_emails": 50, "num_sent": 30, "num_received": 20,
     "first_interaction": "2023-01-01", "last_interaction": "2025-03-15"},
    {"first_name": "Anna", "last_name": "Lee", "company": "Initech", "type": "customer",
     "primary_email": "anna.lee@initech.com", "emails": ["anna.lee@initech.com"],
     "num_emails": 12, "num_sent": 6, "num_received": 6,
     "first_interaction": "2024-02-10", "last_interaction": "2026-01-09"},
    {"first_name": "Sam", "last_name": "Null", "company": "Initech",
     "primary_email": "sam@initech.com", "emails": ["sam@initech.com"],
     "num_emails": 3, "num_sent": 1, "num_received": 2,
     "first_interaction": None, "last_interaction": None},
]


def crit(**kw):
    """Build a criteria dict via the real arg parser for fidelity."""
    argv = ["export", "-i", "x.csv", "-o", "out.csv"]
    for k, v in kw.items():
        argv += ["--" + k.replace("_", "-"), str(v)]
    return build_criteria(parse_args(argv))


def run(**kw):
    c = crit(**kw)
    return [x["first_name"] for x in CONTACTS if matches(x, c)]


class TestTextFilters:
    def test_company_exact_caseinsensitive(self):
        assert run(company="initech") == ["Anna", "Sam"]

    def test_company_set_membership(self):
        assert run(company="Initech,Globex") == ["Jordan", "Anna", "Sam"]

    def test_and_across_columns(self):
        assert run(company="Initech", last_name="Lee") == ["Anna"]

    def test_email_domain(self):
        assert run(email_domain="northwind.com") == ["Dana"]


_BASE = ["export", "-i", "x.csv", "-o", "out.csv"]


class TestTypeFilter:
    def test_type_match(self):
        c = build_criteria(parse_args(_BASE + ["--type", "customer"]))
        assert [x["first_name"] for x in CONTACTS if matches(x, c)] == ["Anna"]

    def test_no_type_flag_matches_all(self):
        c = build_criteria(parse_args(_BASE))
        assert len([x for x in CONTACTS if matches(x, c)]) == len(CONTACTS)

    def test_type_combines_with_other_filters(self):
        c = build_criteria(parse_args(
            _BASE + ["--type", "customer", "--company", "Globex"]))
        # Anna is customer but not Globex; nobody is both -> none.
        assert [x["first_name"] for x in CONTACTS if matches(x, c)] == []


class TestNumericFilters:
    def test_min_inclusive(self):
        assert run(min_emails="12") == ["Jordan", "Anna"]

    def test_range(self):
        assert run(min_emails="5", max_emails="12") == ["Dana", "Anna"]

    def test_min_sent(self):
        assert run(min_sent="6") == ["Jordan", "Anna"]


class TestDateFilters:
    def test_last_after_inclusive(self):
        assert run(last_after="2025-03-15") == ["Jordan", "Anna"]

    def test_last_before(self):
        assert run(last_before="2022-12-31") == ["Dana"]

    def test_first_range(self):
        assert run(first_after="2023-01-01", first_before="2024-12-31") \
            == ["Jordan", "Anna"]

    def test_null_date_excluded_by_filter(self):
        # Sam has null dates and must be excluded once a date bound is active.
        assert "Sam" not in run(last_after="2000-01-01")
