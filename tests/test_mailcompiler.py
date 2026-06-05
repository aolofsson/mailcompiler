"""Unit tests for the mbox contact import helpers."""

import os
import tempfile
from collections import defaultdict

from mailcompiler.mailcompiler import (
    clean, company_from, is_bot, split_name, is_blacklisted, merge_row,
    person_to_row, load_rows, write_rows,
    Rec, _ingest_message, _pst_message_fields, _MAPI_SENDER_SMTP,
    _signature_text, _extract_phones,
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
            "company": "Initech", "phone": "+16502530000",
            "primary_email": "anna@initech.com",
            "emails": ["anna@initech.com"], "num_emails": 10, "num_sent": 4,
            "num_received": 6, "first_interaction": "2023-01-01",
            "last_interaction": "2024-01-01", "source": "alpha.mbox",
        }

    def _new(self):
        return {
            "vip": "", "last_name": "Lee", "first_name": "Anna",
            "company": "Initech", "phone": "+14155552671",
            "primary_email": "anna@initech.com",
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

    def test_phone_preserved(self):
        e = self._existing()                 # has +16502530000
        merge_row(e, self._new())            # new has a different phone
        assert e["phone"] == "+16502530000"  # existing value kept

    def test_phone_filled_when_empty(self):
        e = self._existing()
        e["phone"] = ""
        merge_row(e, self._new())
        assert e["phone"] == "+14155552671"


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
        "company": "Globex", "phone": "+16502530000",
        "primary_email": "jordan@globex.com",
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


def _msg(frm=None, to=None, is_sent=False, is_spam=False, self_hints=None,
         body=""):
    return {"from": frm or [], "to": to or [], "date": None,
            "is_sent": is_sent, "is_spam": is_spam,
            "self_hints": self_hints or [], "body": body}


class TestIngestMessage:
    def test_received_counts_sender(self):
        recs = defaultdict(Rec)
        assert _ingest_message(_msg(frm=[("Bob", "bob@x.com")]), recs, set())
        assert recs["bob@x.com"].num_recv == 1
        assert recs["bob@x.com"].num_sent == 0

    def test_sent_counts_recipients_and_learns_self(self):
        recs, ss = defaultdict(Rec), set()
        _ingest_message(_msg(frm=[("Me", "me@self.com")],
                             to=[("Bob", "bob@x.com")], is_sent=True), recs, ss)
        assert "me@self.com" in ss          # self learned from sent From
        assert recs["bob@x.com"].num_sent == 1
        assert "me@self.com" not in recs    # self never a contact

    def test_spam_skipped(self):
        recs = defaultdict(Rec)
        assert _ingest_message(_msg(frm=[("Bob", "bob@x.com")], is_spam=True),
                               recs, set()) is False
        assert not recs

    def test_self_recipient_skipped(self):
        recs = defaultdict(Rec)
        _ingest_message(_msg(frm=[("Me", "me@self.com")],
                             to=[("Me", "me@self.com"), ("Bob", "bob@x.com")],
                             is_sent=True), recs, {"me@self.com"})
        assert list(recs) == ["bob@x.com"]

    def test_self_sender_not_added(self):
        recs = defaultdict(Rec)
        # From is self -> treated as sent; the (empty) recipient list adds nobody.
        _ingest_message(_msg(frm=[("Me", "me@self.com")]), recs, {"me@self.com"})
        assert not recs

    def test_received_signature_phone_recorded(self):
        recs = defaultdict(Rec)
        body = "Thanks,\nBob\n--\nBob Jones\nMobile: (650) 253-0000\n"
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")], body=body), recs, set())
        assert recs["bob@x.com"].phones.get("+16502530000") == 1

    def test_sent_signature_phone_ignored(self):
        recs, ss = defaultdict(Rec), set()
        body = "Regards,\nMe\n--\nMe\nMobile: (650) 253-0000\n"
        _ingest_message(_msg(frm=[("Me", "me@self.com")],
                             to=[("Bob", "bob@x.com")], is_sent=True,
                             body=body), recs, ss)
        assert not recs["bob@x.com"].phones   # our own signature is not theirs


class TestSignatureText:
    def test_after_dash_delimiter(self):
        body = "Hi there\nbody line\n--\nAlice\nTel: 1\n"
        sig = _signature_text(body)
        assert "Alice" in sig and "body line" not in sig

    def test_quoted_reply_excluded(self):
        body = "Thanks!\n\nOn Mon, X wrote:\n> call (650) 253-0000\n"
        assert "650" not in _signature_text(body)


class TestExtractPhones:
    def test_labeled_us_number(self):
        body = "Best,\nBob\n--\nBob\nMobile: (650) 253-0000\n"
        assert _extract_phones(body) == ["+16502530000"]

    def test_explicit_international(self):
        body = "--\nDana\nTel: +44 20 7031 3000\n"
        assert _extract_phones(body) == ["+442070313000"]

    def test_number_only_in_quote_ignored(self):
        body = "Cheers\n\nOn Tue someone wrote:\n> reach me at (650) 253-0000\n"
        assert _extract_phones(body) == []

    def test_non_phone_not_matched(self):
        body = "--\nOrder #1234567 shipped on 2024-01-02\n"
        assert _extract_phones(body) == []


class _Entry:
    def __init__(self, t, v):
        self.entry_type, self._v = t, v

    def get_data_as_string(self):
        return self._v


class _RS:
    def __init__(self, entries):
        self._e = entries
        self.number_of_entries = len(entries)

    def get_entry(self, i):
        return self._e[i]


class _PstMsg:
    def __init__(self, headers=None, sender_name=None, record_sets=None):
        self.transport_headers = headers
        self.sender_name = sender_name
        self._rs = record_sets or []
        self.number_of_record_sets = len(self._rs)
        self.client_submit_time = None
        self.delivery_time = None

    def get_record_set(self, i):
        return self._rs[i]


class TestPstMessageFields:
    HEADERS = ("From: Bob <bob@x.com>\nTo: Me <me@self.com>\n"
               "Cc: C <c@y.com>\nDate: Wed, 03 Jun 2026 14:35:06 +0000\n\n")

    def test_transport_headers_path(self):
        f = _pst_message_fields(_PstMsg(headers=self.HEADERS), "Inbox")
        assert f["from"] == [("Bob", "bob@x.com")]
        assert ("Me", "me@self.com") in f["to"]
        assert ("C", "c@y.com") in f["to"]
        assert f["is_sent"] is False and f["is_spam"] is False
        assert f["date"] is not None

    def test_sent_and_junk_folders(self):
        assert _pst_message_fields(_PstMsg(headers="From: a@b.com\n\n"),
                                   "Sent Items")["is_sent"] is True
        assert _pst_message_fields(_PstMsg(headers="From: a@b.com\n\n"),
                                   "Junk Email")["is_spam"] is True

    def test_mapi_fallback_sender(self):
        m = _PstMsg(headers=None, sender_name="Carol",
                    record_sets=[_RS([_Entry(_MAPI_SENDER_SMTP, "carol@z.com")])])
        f = _pst_message_fields(m, "Inbox")
        assert f["from"] == [("Carol", "carol@z.com")]
        assert f["to"] == []

    def test_mapi_fallback_nonsmtp_dropped(self):
        m = _PstMsg(headers=None, sender_name="Carol",
                    record_sets=[_RS([_Entry(_MAPI_SENDER_SMTP, "/O=EX/CN=carol")])])
        f = _pst_message_fields(m, "Inbox")
        assert f["from"] == [("Carol", "")]   # non-SMTP -> blanked, later skipped
