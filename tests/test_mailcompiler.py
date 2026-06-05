"""Unit tests for the mbox contact import helpers."""

import os
import tempfile

from mailcompiler.mailcompiler import (
    clean, company_from, is_bot, split_name, is_blacklisted, merge_row,
    person_to_row, load_rows, write_rows,
)


class TestCompanyFrom:
    def test_free_provider_is_blank(self):
        assert company_from("someone@gmail.com") == ""
        assert company_from("someone@yahoo.co.uk") == ""

    def test_corporate_domain(self):
        assert company_from("sam@acme.com") == "Acme"
        assert company_from("a@intel.com") == "Intel"

    def test_subdomain_uses_registrable_part(self):
        assert company_from("yvonne@us.bosch.com") == "Bosch"


class TestIsBot:
    def test_human_addresses_are_not_bots(self):
        assert not is_bot("sam@acme.com")
        assert not is_bot("chris.morgan@us.af.mil")

    def test_noreply_variants(self):
        assert is_bot("no-reply@example.com")
        assert is_bot("do_not_reply@linklings.com")
        assert is_bot("noreply@example.com")

    def test_tagged_and_github_replies(self):
        assert is_bot("reply+deadbeef@reply.github.com")
        assert is_bot("notifications@github.com")

    def test_bulk_domains_and_substrings(self):
        assert is_bot("marketingemail@3dsystems.com")
        assert is_bot("hash.1.2.3@unsub-sj.mktomail.com")

    def test_missing_domain(self):
        assert is_bot("brokenaddress")


class TestSplitName:
    def test_first_last(self):
        assert split_name("Pat Carter", "x@y.com") == ("Pat", "Carter")

    def test_last_comma_first(self):
        assert split_name("Lopez, Maria", "x@y.com") == ("Maria", "Lopez")

    def test_strips_quotes_and_dept_suffix(self):
        # leading quote and a department suffix after a paren/slash are dropped
        first, last = split_name("Tan (Alex)/dept", "alex.tan@sk.com")
        assert first == "Tan"

    def test_fallback_from_email_localpart(self):
        assert split_name("", "chris.morgan.4@us.af.mil") == ("Chris", "Morgan")

    def test_fallback_rejects_hash_localpart(self):
        assert split_name("", "8a795fa6-038a-4166@unsub.beehiiv.com") == ("", "")

    def test_single_token_has_no_last_name(self):
        first, last = split_name("Robin", "robin@acme.com")
        assert first == "Robin"
        assert last == ""


class TestClean:
    def test_strips_wrapping_punctuation(self):
        assert clean("'Bank") == "Bank"
        assert clean('"Name".') == "Name"
        assert clean("  spaced  ") == "spaced"


class TestIsBlacklisted:
    def test_exact_domain(self):
        assert is_blacklisted("a@blocked.example", {"blocked.example"})

    def test_subdomain(self):
        assert is_blacklisted("a@sub.blocked.example", {"blocked.example"})

    def test_not_blacklisted(self):
        assert not is_blacklisted("sam@acme.com", {"blocked.example"})

    def test_empty_blacklist(self):
        assert not is_blacklisted("a@b.com", set())


class TestMergeRow:
    def _existing(self):
        return {
            "vip": "Y", "last_name": "Lee", "first_name": "Anna",
            "company": "Initech", "primary_email": "anna@initech.com",
            "emails": ["anna@initech.com"], "num_emails": 10, "num_sent": 4,
            "num_received": 6, "first_interaction": "2023-01-01",
            "last_interaction": "2024-01-01", "source": "alpha.mbox",
        }

    def _new(self):
        return {
            "vip": "", "last_name": "Lee", "first_name": "Anna",
            "company": "Initech", "primary_email": "anna@initech.com",
            "emails": ["anna@initech.com", "anna.lee@initech.com"],
            "num_emails": 5, "num_sent": 2, "num_received": 3,
            "first_interaction": "2022-06-01", "last_interaction": "2025-03-15",
            "source": "beta.mbox",
        }

    def test_counts_are_overwritten(self):
        e = self._existing()
        merge_row(e, self._new())
        assert e["num_emails"] == 5
        assert e["num_sent"] == 2
        assert e["num_received"] == 3

    def test_vip_and_text_preserved(self):
        e = self._existing()
        e["vip"] = "Y"
        merge_row(e, self._new())
        assert e["vip"] == "Y"  # hand-edited field not clobbered

    def test_emails_union_and_dates_widen(self):
        e = self._existing()
        merge_row(e, self._new())
        assert e["emails"] == ["anna@initech.com", "anna.lee@initech.com"]
        assert e["first_interaction"] == "2022-06-01"  # earliest
        assert e["last_interaction"] == "2025-03-15"    # latest

    def test_empty_existing_text_filled_from_new(self):
        e = self._existing()
        e["company"] = ""
        merge_row(e, self._new())
        assert e["company"] == "Initech"

    def test_sources_union(self):
        e = self._existing()
        merge_row(e, self._new())
        assert e["source"] == "alpha.mbox | beta.mbox"
        # re-merging the same source does not duplicate it
        merge_row(e, self._new())
        assert e["source"] == "alpha.mbox | beta.mbox"


class TestPersonToRow:
    def test_tags_source(self):
        p = {"first_name": "A", "last_name": "B", "company": "C",
             "primary_email": "a@b.com", "emails": ["a@b.com"],
             "num_emails": 1, "num_sent": 1, "num_received": 0,
             "first_interaction": "2024-01-01", "last_interaction": "2024-01-01"}
        row = person_to_row(p, "All mail.mbox")
        assert row["source"] == "All mail.mbox"
        assert row["vip"] == ""


class TestRoundTrip:
    ROWS = [{
        "vip": "Y", "last_name": "Vale", "first_name": "Jordan",
        "company": "Globex", "primary_email": "jordan@globex.com",
        "emails": ["jordan@globex.com", "jordan.vale@globex.com"],
        "num_emails": 3, "num_sent": 2, "num_received": 1,
        "first_interaction": "2023-01-01", "last_interaction": "2024-05-05",
        "source": "a.mbox | b.mbox",
    }]

    def _roundtrip(self, ext):
        fd, path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        try:
            write_rows(path, self.ROWS)
            back = load_rows(path)
        finally:
            os.remove(path)
        return back

    def test_json_roundtrip(self):
        back = self._roundtrip(".json")
        assert back == [
            {k: self.ROWS[0][k] for k in self.ROWS[0]}
        ]

    def test_csv_roundtrip(self):
        back = self._roundtrip(".csv")
        r = back[0]
        # CSV preserves the same values (emails re-split, ints re-parsed)
        assert r["emails"] == ["jordan@globex.com", "jordan.vale@globex.com"]
        assert r["num_sent"] == 2 and r["num_received"] == 1
        assert r["vip"] == "Y" and r["source"] == "a.mbox | b.mbox"
        assert r["last_interaction"] == "2024-05-05"
