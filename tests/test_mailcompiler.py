"""Unit tests for the mbox contact import helpers."""

import csv
import json
import os
import tempfile
from collections import defaultdict

from mailcompiler.mailcompiler import (
    clean, company_from, is_bot, split_name, is_blacklisted, merge_row,
    person_to_row, load_rows, write_contacts_as,
    Rec, _ingest_message, _pst_message_fields, _MAPI_SENDER_SMTP,
    _signature_text, _extract_phones,
    _format_addrs, _is_noreply, dump_llm,
    _join_distinct, _merge_group, dedup_contacts,
    _vcard_escape, _vcard_fold, _contact_vcard, write_vcards,
    _vcard_unescape, parse_vcards,
    OUTLOOK_FIELDS, write_outlook_csv, parse_outlook_csv,
    write_xlsx_rows, write_outlook_xlsx, parse_outlook_xlsx,
    _write_xlsx, _read_xlsx, iter_mbox_messages,
    load_domain_list, contact_domains, contact_in_domains, select_by_domains,
    main,
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

    def test_punctuation_only_display_name(self):
        # A display name that cleans to nothing must not crash (regression).
        # Numeric local-part so the email fallback also yields no name.
        for junk in (".", ". .", "''", "()", " , ", "-.-"):
            assert split_name(junk, "12345@x.com") == ("", "")


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
            "type": "customer", "friend": "Y", "last_name": "Lee",
            "first_name": "Anna",
            "title": "VP", "company": "Initech", "phone": "+16502530000",
            "address": "1 Main St", "primary_email": "anna@initech.com",
            "emails": ["anna@initech.com"], "num_emails": 10, "num_sent": 4,
            "num_received": 6, "first_interaction": "2023-01-01",
            "last_interaction": "2024-01-01", "source": "alpha.mbox",
        }

    def _new(self):
        return {
            "type": "", "friend": "", "last_name": "Lee", "first_name": "Anna",
            "title": "Director", "company": "Initech", "phone": "+14155552671",
            "address": "", "primary_email": "anna@initech.com",
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

    def test_type_and_text_preserved(self):
        e = self._existing()
        merge_row(e, self._new())
        assert e["type"] == "customer"   # hand-edited field not clobbered
        assert e["title"] == "VP"        # existing title kept over the new one

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
        assert row["type"] == ""


class TestRoundTrip:
    ROWS = [{
        "type": "customer", "friend": "Y", "last_name": "Vale",
        "first_name": "Jordan",
        "title": "CTO", "company": "Globex", "phone": "+16502530000",
        "address": "10 Loop, CA", "primary_email": "jordan@globex.com",
        "emails": ["jordan@globex.com", "jordan.vale@globex.com"],
        "num_emails": 3, "num_sent": 2, "num_received": 1, "num_cc": 0,
        "first_interaction": "2023-01-01", "last_interaction": "2024-05-05",
        "source": "a.mbox | b.mbox",
    }]

    def _roundtrip(self, ext):
        fd, path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        fmt = {".json": "json", ".csv": "csv", ".xlsx": "xlsx"}[ext]
        try:
            write_contacts_as(path, self.ROWS, fmt)
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
        assert r["type"] == "customer" and r["source"] == "a.mbox | b.mbox"
        assert r["last_interaction"] == "2024-05-05"

    def test_xlsx_roundtrip(self):
        back = self._roundtrip(".xlsx")
        r = back[0]
        assert r["emails"] == ["jordan@globex.com", "jordan.vale@globex.com"]
        assert r["num_sent"] == 2 and r["num_received"] == 1
        assert r["type"] == "customer" and r["source"] == "a.mbox | b.mbox"

    def test_json_csv_xlsx_all_equal(self):
        # The three native formats carry identical information.
        assert self._roundtrip(".json") == self._roundtrip(".csv") \
            == self._roundtrip(".xlsx")


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

    def test_received_corecipients_ignored_by_default(self):
        recs = defaultdict(Rec)
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")],
                             to=[("Me", "me@self.com"), ("Sue", "sue@y.com")]),
                        recs, {"me@self.com"})
        assert "bob@x.com" in recs and recs["bob@x.com"].num_recv == 1
        assert "sue@y.com" not in recs        # co-recipient not harvested

    def test_include_cc_harvests_corecipients(self):
        recs = defaultdict(Rec)
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")],
                             to=[("Me", "me@self.com"), ("Sue", "sue@y.com")]),
                        recs, {"me@self.com"}, include_cc=True)
        assert recs["bob@x.com"].num_recv == 1   # sender still num_recv
        assert recs["bob@x.com"].num_cc == 0
        assert recs["sue@y.com"].num_cc == 1     # co-recipient counted as cc
        assert recs["sue@y.com"].num_recv == 0
        assert "me@self.com" not in recs         # self never harvested

    def test_include_cc_does_not_double_count_sender(self):
        recs = defaultdict(Rec)
        # sender also appears in To: should not get both num_recv and num_cc
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")],
                             to=[("Bob", "bob@x.com"), ("Me", "me@self.com")]),
                        recs, {"me@self.com"}, include_cc=True)
        assert recs["bob@x.com"].num_recv == 1
        assert recs["bob@x.com"].num_cc == 0


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


def _llm_msg(subject="", frm=None, to=None, date=None, body="", is_spam=False):
    return {"subject": subject, "from": frm or [], "to": to or [],
            "date": date, "is_sent": False, "is_spam": is_spam,
            "self_hints": [], "body": body}


class TestLlmCorpus:
    def test_format_addrs(self):
        pairs = [("Bob", "bob@x.com"), ("", "c@y.com")]
        assert _format_addrs(pairs) == "Bob <bob@x.com>, c@y.com"

    def test_format_addrs_decodes_name(self):
        pairs = [("=?utf-8?q?Jos=C3=A9?=", "jose@x.com")]
        assert _format_addrs(pairs) == "José <jose@x.com>"

    def test_is_noreply(self):
        assert _is_noreply([("Svc", "no-reply@svc.com")])
        assert _is_noreply([("noreply", "x@y.com")])
        assert not _is_noreply([("Bob", "bob@x.com")])

    def test_dump_llm_jsonl(self):
        from datetime import datetime
        msgs = [
            _llm_msg("Hi", [("Bob", "bob@x.com")], [("Me", "me@self.com")],
                     None, "hello"),
            _llm_msg("Promo", [("N", "no-reply@svc.com")], [], None, "buy"),
            _llm_msg("Re: Hi", [("Ann", "ann@y.com")], [],
                     datetime(2024, 5, 5, 9, 0, 0), "thanks"),
        ]
        fd, path = tempfile.mkstemp(suffix=".jsonl")
        os.close(fd)
        try:
            n = dump_llm(iter(msgs), path)
            lines = [json.loads(ln) for ln in open(path) if ln.strip()]
        finally:
            os.remove(path)
        assert n == 2                                   # no-reply skipped
        assert {ln["subject"] for ln in lines} == {"Hi", "Re: Hi"}
        first = lines[0]
        assert set(first) == {"subject", "from", "to", "date", "body"}
        assert first["from"] == "Bob <bob@x.com>"
        assert first["body"] == "hello"
        assert lines[1]["date"] == "2024-05-05T09:00:00"

    def test_dump_llm_filters_and_strips(self):
        msgs = [
            _llm_msg("a", [("Bob", "bob@x.com")],
                     body="Hi there\nThanks\nOn Mon, X wrote:\n> old stuff"),
            _llm_msg("spam", [("S", "s@x.com")], body="buy", is_spam=True),
            _llm_msg("bot", [("GH", "notifications@github.com")], body="ping"),
            _llm_msg("empty", [("E", "e@x.com")], body="   "),
            _llm_msg("d1", [("D", "d@x.com")], body="same body"),
            _llm_msg("d2", [("E2", "e2@x.com")], body="same body"),
        ]
        fd, path = tempfile.mkstemp(suffix=".jsonl")
        os.close(fd)
        try:
            n = dump_llm(iter(msgs), path)
            lines = [json.loads(ln) for ln in open(path) if ln.strip()]
        finally:
            os.remove(path)
        assert n == 2                       # spam, bot, empty, duplicate dropped
        bob = [ln for ln in lines if ln["subject"] == "a"][0]
        assert bob["body"] == "Hi there\nThanks"          # quoted history stripped
        assert sum(ln["body"] == "same body" for ln in lines) == 1   # deduped


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


def _row(first="", last="", company="", phone="", ctype="", primary="",
         emails=None, n_emails=0, n_sent=0, n_recv=0, first_i=None, last_i=None,
         source="", title="", address="", friend=""):
    return {"type": ctype, "friend": friend, "last_name": last,
            "first_name": first, "title": title, "company": company,
            "phone": phone, "address": address, "primary_email": primary,
            "emails": emails if emails is not None else ([primary] if primary else []),
            "num_emails": n_emails, "num_sent": n_sent, "num_received": n_recv,
            "first_interaction": first_i, "last_interaction": last_i,
            "source": source}


class TestJoinDistinct:
    def test_dedupes_and_drops_blanks(self):
        assert _join_distinct(["Acme", "", "Acme", "Globex"]) == "Acme | Globex"

    def test_splits_prejoined(self):
        assert _join_distinct(["A | B", "B | C"]) == "A | B | C"


class TestMergeGroup:
    def _pair(self):
        r1 = _row("Jane", "Roe", "Acme", "+16175550000", "customer",
                  "jane@acme.com", ["jane@acme.com"], 40, 30, 10,
                  "2022-01-01", "2024-01-01", "a.mbox", title="VP",
                  address="1 Main", friend="Y")
        r2 = _row("Jane", "Roe", "Acme2", "+14155550000", "competitor",
                  "jane@gmail.com", ["jane@gmail.com"], 3, 2, 1,
                  "2021-06-01", "2025-05-05", "b.mbox", title="Director",
                  address="2 Oak", friend="")
        return r1, r2

    def test_counts_summed(self):
        m = _merge_group(list(self._pair()))
        assert (m["num_emails"], m["num_sent"], m["num_received"]) == (43, 32, 11)

    def test_emails_union_and_winner_fields(self):
        m = _merge_group(list(self._pair()))
        assert m["emails"] == ["jane@acme.com", "jane@gmail.com"]
        assert m["primary_email"] == "jane@acme.com"   # higher-volume record

    def test_conflicts_joined_and_flags_kept(self):
        m = _merge_group(list(self._pair()))
        assert m["company"] == "Acme | Acme2"
        assert m["phone"] == "+16175550000 | +14155550000"
        assert m["type"] == "customer | competitor"
        assert m["friend"] == "Y"
        assert m["title"] == "VP | Director"
        assert m["address"] == "1 Main | 2 Oak"
        assert m["source"] == "a.mbox | b.mbox"

    def test_dates_widen(self):
        m = _merge_group(list(self._pair()))
        assert m["first_interaction"] == "2021-06-01"
        assert m["last_interaction"] == "2025-05-05"


class TestDedupContacts:
    def test_merges_case_insensitive_and_keeps_others(self):
        rows = [
            _row("Jane", "Roe", "Acme", primary="jane@acme.com", n_emails=40),
            _row("jane", "roe", "Acme2", primary="jane@gmail.com", n_emails=3),
            _row("Bob", "Smith", "Globex", primary="bob@globex.com", n_emails=5),
            _row("Support", "", "Acme", primary="support@acme.com", n_emails=9),
        ]
        out = dedup_contacts(rows)
        assert len(out) == 3                       # Jane*2 merged; Bob; Support
        janes = [r for r in out if r["last_name"].lower() == "roe"]
        assert len(janes) == 1 and janes[0]["num_emails"] == 43
        assert janes[0]["first_name"] == "Jane"    # casing from higher-volume row
        # blank-last-name row passes through untouched
        assert any(r["first_name"] == "Support" and r["last_name"] == ""
                   for r in out)

    def test_idempotent(self):
        rows = [_row("Bob", "Smith", primary="bob@x.com", n_emails=5)]
        assert len(dedup_contacts(dedup_contacts(rows))) == 1


class TestVcard:
    def test_escape(self):
        assert _vcard_escape("a,b;c\\d\ne") == "a\\,b\\;c\\\\d\\ne"

    def test_fold_ascii(self):
        folded = _vcard_fold("NOTE:" + "x" * 200)
        physical = folded.split("\r\n")
        assert len(physical) > 1
        assert all(len(p.encode("utf-8")) <= 75 for p in physical)
        assert all(p.startswith(" ") for p in physical[1:])  # continuations
        assert folded.replace("\r\n ", "") == "NOTE:" + "x" * 200

    def test_fold_multibyte_safe(self):
        line = "NOTE:" + "é" * 100        # 2-byte chars
        folded = _vcard_fold(line)
        assert all(len(p.encode("utf-8")) <= 75 for p in folded.split("\r\n"))
        assert folded.replace("\r\n ", "") == line   # no char was split

    def test_contact_vcard_fields(self):
        row = _row("Jane", "Roe", "Acme", "+16175550000", "customer",
                   "jane@acme.com", ["jane@acme.com", "jane@gmail.com"],
                   40, 30, 10, "2022-01-01", "2024-01-01", "a.mbox",
                   title="VP", address="1 Main St", friend="Y")
        lines = _contact_vcard(row)
        assert lines[0] == "BEGIN:VCARD" and lines[-1] == "END:VCARD"
        assert "VERSION:3.0" in lines
        assert "FN:Jane Roe" in lines
        assert "N:Roe;Jane;;;" in lines
        assert "ORG:Acme" in lines and "TITLE:VP" in lines
        assert "EMAIL;TYPE=INTERNET,PREF:jane@acme.com" in lines
        assert "EMAIL;TYPE=INTERNET:jane@gmail.com" in lines
        assert "TEL;TYPE=VOICE:+16175550000" in lines
        assert "CATEGORIES:customer,friend" in lines
        assert any(ln.startswith("NOTE:") for ln in lines)

    def test_write_vcards_file(self):
        rows = [_row("Jane", "Roe", primary="jane@acme.com", n_emails=1),
                _row("Bob", "Smith", primary="bob@x.com", n_emails=1)]
        fd, path = tempfile.mkstemp(suffix=".vcf")
        os.close(fd)
        try:
            write_vcards(path, rows)
            data = open(path, "rb").read()
        finally:
            os.remove(path)
        assert data.count(b"BEGIN:VCARD") == 2
        assert b"\r\n" in data
        assert all(len(p) <= 75 for p in data.split(b"\r\n"))


GMAIL_VCF = """BEGIN:VCARD
VERSION:3.0
FN:Alex Smith
N:Smith;Alex;;;
ORG:Acme Corp;Sales
TITLE:Account Executive
EMAIL;TYPE=INTERNET,WORK,PREF:alex.smith@acme.example.com
EMAIL;TYPE=INTERNET,HOME:alex.personal@example.com
TEL;TYPE=CELL,VOICE,PREF:+1-555-0199
ADR;TYPE=WORK,PREF:;;123 Business Rd;Boston;MA;02110;USA
NOTE:Met at the 2026 tech conference.
CATEGORIES:customer,friend
END:VCARD
"""


def _write_tmp(text, suffix):
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "w", newline="") as fh:
        fh.write(text)
    return path


class TestVcardImport:
    def test_unescape(self):
        assert _vcard_unescape("a\\,b\\;c\\nd\\\\e") == "a,b;c\nd\\e"

    def test_parse_gmail_reference(self):
        path = _write_tmp(GMAIL_VCF, ".vcf")
        try:
            rows = parse_vcards(path)
        finally:
            os.remove(path)
        assert len(rows) == 1
        r = rows[0]
        assert (r["first_name"], r["last_name"]) == ("Alex", "Smith")
        assert r["company"] == "Acme Corp"          # first ORG component
        assert r["title"] == "Account Executive"
        assert r["primary_email"] == "alex.smith@acme.example.com"   # PREF
        assert "alex.personal@example.com" in r["emails"]
        assert r["phone"] == "+1-555-0199"
        assert "Boston" in r["address"]
        assert r["type"] == "customer" and r["friend"] == "Y"

    def test_roundtrip_write_then_parse(self):
        rows = [_row("Jane", "Roe", "Acme", "+16175550000", "customer",
                     "jane@acme.com", ["jane@acme.com", "jane@gmail.com"],
                     title="VP", address="1 Main St", friend="Y")]
        path = _write_tmp("", ".vcf")
        try:
            write_vcards(path, rows)
            back = parse_vcards(path)
        finally:
            os.remove(path)
        assert len(back) == 1
        b = back[0]
        assert (b["first_name"], b["last_name"]) == ("Jane", "Roe")
        assert b["primary_email"] == "jane@acme.com"
        assert b["emails"] == ["jane@acme.com", "jane@gmail.com"]
        assert b["company"] == "Acme" and b["title"] == "VP"
        assert b["phone"] == "+16175550000"
        assert b["type"] == "customer" and b["friend"] == "Y"


OUTLOOK_VCF_CSV = (
    "First Name,Last Name,Job Title,Company,E-mail Address,E-mail 2 Address,"
    "Business Phone,Business Street,Business City,Business State\r\n"
    "Alex,Smith,Account Executive,Acme Corp,alex@acme.com,alex2@acme.com,"
    "+1-555-0199,123 Business Rd,Boston,MA\r\n")


class TestOutlookCsv:
    def test_write_outlook_csv(self):
        rows = [_row("Jane", "Roe", "Acme", "+16175550000", "customer",
                     "jane@acme.com", ["jane@acme.com", "jane@gmail.com"],
                     title="VP", address="1 Main St")]
        path = _write_tmp("", ".csv")
        try:
            write_outlook_csv(path, rows)
            with open(path, newline="") as fh:
                rd = list(csv.reader(fh))
        finally:
            os.remove(path)
        assert rd[0] == OUTLOOK_FIELDS
        rec = dict(zip(rd[0], rd[1]))
        assert rec["First Name"] == "Jane" and rec["Last Name"] == "Roe"
        assert rec["Job Title"] == "VP" and rec["Company"] == "Acme"
        assert rec["E-mail Address"] == "jane@acme.com"
        assert rec["E-mail 2 Address"] == "jane@gmail.com"
        assert rec["Business Phone"] == "+16175550000"
        assert rec["Business Street"] == "1 Main St"

    def test_parse_outlook_csv(self):
        path = _write_tmp(OUTLOOK_VCF_CSV, ".csv")
        try:
            rows = parse_outlook_csv(path)
        finally:
            os.remove(path)
        assert len(rows) == 1
        r = rows[0]
        assert (r["first_name"], r["last_name"]) == ("Alex", "Smith")
        assert r["title"] == "Account Executive" and r["company"] == "Acme Corp"
        assert r["phone"] == "+1-555-0199"
        assert r["emails"] == ["alex@acme.com", "alex2@acme.com"]
        assert r["address"] == "123 Business Rd, Boston, MA"
        assert r["type"] == "" and r["friend"] == "" and r["num_emails"] == 0

    def test_roundtrip(self):
        rows = [_row("Jane", "Roe", "Acme", "+16175550000", "customer",
                     "jane@acme.com", ["jane@acme.com", "jane@gmail.com"],
                     title="VP", address="1 Main St")]
        path = _write_tmp("", ".csv")
        try:
            write_outlook_csv(path, rows)
            back = parse_outlook_csv(path)
        finally:
            os.remove(path)
        b = back[0]
        assert (b["first_name"], b["last_name"]) == ("Jane", "Roe")
        assert b["company"] == "Acme" and b["title"] == "VP"
        assert b["phone"] == "+16175550000"
        assert b["emails"] == ["jane@acme.com", "jane@gmail.com"]
        assert b["address"] == "1 Main St"
        assert b["type"] == "" and b["num_emails"] == 0   # not carried by Outlook


class TestXlsx:
    def test_write_read_roundtrip(self):
        header = ["A", "B", "C"]
        rows = [["1", "x,y", ""], ["2", "z", "q"]]
        path = _write_tmp("", ".xlsx")
        try:
            _write_xlsx(path, header, rows)
            back = _read_xlsx(path)
        finally:
            os.remove(path)
        assert back == [{"A": "1", "B": "x,y", "C": ""},
                        {"A": "2", "B": "z", "C": "q"}]

    def test_native_xlsx_load(self):
        rows = [_row("Jane", "Roe", "Acme", primary="jane@acme.com",
                     emails=["jane@acme.com", "j2@acme.com"], n_emails=5)]
        path = _write_tmp("", ".xlsx")
        try:
            write_xlsx_rows(path, rows)
            back = load_rows(path)
        finally:
            os.remove(path)
        r = back[0]
        assert (r["first_name"], r["last_name"]) == ("Jane", "Roe")
        assert r["emails"] == ["jane@acme.com", "j2@acme.com"]
        assert r["num_emails"] == 5

    def test_outlook_xlsx_roundtrip(self):
        rows = [_row("Alex", "Smith", "Acme", "+15550199", "customer",
                     "alex@acme.com", ["alex@acme.com", "alex2@acme.com"],
                     title="AE", address="123 Business Rd")]
        path = _write_tmp("", ".xlsx")
        try:
            write_outlook_xlsx(path, rows)
            back = parse_outlook_xlsx(path)
        finally:
            os.remove(path)
        b = back[0]
        assert (b["first_name"], b["last_name"]) == ("Alex", "Smith")
        assert b["company"] == "Acme" and b["title"] == "AE"
        assert b["phone"] == "+15550199"
        assert b["emails"] == ["alex@acme.com", "alex2@acme.com"]
        assert b["address"] == "123 Business Rd"


class TestBodyCap:
    def _mbox(self, body):
        return ("From 100@xxx Wed Jun 03 14:35:08 +0000 2026\n"
                "From: Bob <bob@x.com>\n"
                "To: me@self.com\n"
                "Date: Wed, 03 Jun 2026 14:35:06 +0000\n"
                'Content-Type: text/plain; charset="utf-8"\n'
                "\n" + body + "\n")

    def test_body_is_capped(self):
        path = _write_tmp(self._mbox("x" * 5000), ".mbox")
        try:
            capped = list(iter_mbox_messages(path, body_cap=400))[0]["body"]
            full = list(iter_mbox_messages(path, body_cap=None))[0]["body"]
        finally:
            os.remove(path)
        assert 0 < len(capped) <= 400      # cap honored
        assert len(full) >= 5000           # uncapped keeps the whole body
        assert len(capped) < len(full)


class TestLoadDomainList:
    def test_skips_comments_and_blanks(self):
        text = ("# header\n"
                "intel.com\n"
                "\n"
                "   # indented comment\n"
                "@nvidia.com\n"
                "  GD.com  \n")
        path = _write_tmp(text, ".txt")
        try:
            domains = load_domain_list(path)
        finally:
            os.remove(path)
        assert domains == {"intel.com", "nvidia.com", "gd.com"}


class TestContactDomains:
    def test_collects_primary_and_emails(self):
        c = _row(primary="a@intel.com", emails=["a@intel.com", "a@fab.intel.de"])
        assert contact_domains(c) == {"intel.com", "fab.intel.de"}

    def test_in_domains_exact_and_subdomain(self):
        c = _row(primary="bob@fab.intel.com")
        assert contact_in_domains(c, {"intel.com"})        # subdomain matches
        assert contact_in_domains(_row(primary="x@gd.com"), {"gd.com"})
        assert not contact_in_domains(_row(primary="x@intel.de"), {"intel.com"})

    def test_matches_any_email(self):
        c = _row(primary="p@gmail.com", emails=["p@gmail.com", "p@nvidia.com"])
        assert contact_in_domains(c, {"nvidia.com"})       # secondary matches


class TestSelectByDomains:
    def _people(self):
        return [
            _row(primary="a@intel.com"),
            _row(primary="b@nvidia.com"),
            _row(primary="c@gmail.com"),
            _row(primary="d@sub.intel.com"),
        ]

    def test_no_filters_is_identity(self):
        kept, nw, nb = select_by_domains(self._people(), None, None)
        assert len(kept) == 4 and nw == 0 and nb == 0

    def test_whitelist_keeps_only_matches(self):
        kept, nw, nb = select_by_domains(self._people(), {"intel.com"}, None)
        emails = {c["primary_email"] for c in kept}
        assert emails == {"a@intel.com", "d@sub.intel.com"}   # subdomain too
        assert nw == 2 and nb == 0

    def test_blacklist_drops_matches(self):
        kept, nw, nb = select_by_domains(self._people(), None, {"gmail.com"})
        assert "c@gmail.com" not in {c["primary_email"] for c in kept}
        assert nw == 0 and nb == 1

    def test_whitelist_then_blacklist(self):
        kept, nw, nb = select_by_domains(
            self._people(), {"intel.com", "nvidia.com"}, {"nvidia.com"})
        assert {c["primary_email"] for c in kept} == {
            "a@intel.com", "d@sub.intel.com"}
        assert nw == 1 and nb == 1


class TestWhitelistExportCli:
    def test_export_filters_by_whitelist(self):
        rows = [
            _row(first="A", last="One", primary="a@intel.com", n_recv=1),
            _row(first="B", last="Two", primary="b@gmail.com", n_recv=1),
            _row(first="C", last="Three", primary="c@nvidia.com", n_recv=1),
        ]
        dbp = _write_tmp(json.dumps(rows), ".json")
        wlp = _write_tmp("# semis\nintel.com\nnvidia.com\n", ".txt")
        outp = _write_tmp("", ".csv")
        try:
            main(["-i", dbp, "-o", outp, "--whitelist", wlp])
            got = load_rows(outp)
        finally:
            for p in (dbp, wlp, outp):
                os.remove(p)
        emails = {c["primary_email"] for c in got}
        assert emails == {"a@intel.com", "c@nvidia.com"}  # gmail dropped

    def test_export_unions_multiple_whitelist_files(self):
        rows = [
            _row(first="A", last="One", primary="a@intel.com", n_recv=1),
            _row(first="B", last="Two", primary="b@gmail.com", n_recv=1),
            _row(first="C", last="Three", primary="c@lockheedmartin.com", n_recv=1),
        ]
        dbp = _write_tmp(json.dumps(rows), ".json")
        semis = _write_tmp("# semis\nintel.com\n", ".txt")
        defense = _write_tmp("# defense\nlockheedmartin.com\n", ".txt")
        outp = _write_tmp("", ".csv")
        try:
            # two files passed to one flag are unioned
            main(["-i", dbp, "-o", outp, "--whitelist", semis, defense])
            got = load_rows(outp)
        finally:
            for p in (dbp, semis, defense, outp):
                os.remove(p)
        emails = {c["primary_email"] for c in got}
        assert emails == {"a@intel.com", "c@lockheedmartin.com"}
