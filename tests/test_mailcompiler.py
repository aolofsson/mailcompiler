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
    parse_linkedin_csv, _name_key, _fold_linkedin,
    reconcile_contacts,
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

    def test_force_overwrites_text(self):
        e = self._existing()
        merge_row(e, self._new(), force=True)
        assert e["type"] == "customer"      # new is blank -> existing kept
        assert e["title"] == "Director"     # new non-empty -> overwrites
        assert e["phone"] == "+14155552671"  # new non-empty -> overwrites
        assert e["emails"] == ["anna@initech.com", "anna.lee@initech.com"]


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
        "num_emails": 3, "num_sent": 2, "num_received": 1,
        "first_interaction": "2023-01-01", "last_interaction": "2024-05-05",
        "source": "a.mbox | b.mbox",
        "linkedin": "https://www.linkedin.com/in/jordanvale", "import_date": "2026-06-06",
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

    def test_corecipients_harvested_by_default(self):
        recs = defaultdict(Rec)
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")],
                             to=[("Me", "me@self.com"), ("Sue", "sue@y.com")]),
                        recs, {"me@self.com"})
        assert recs["bob@x.com"].num_recv == 1   # sender counts as received
        assert "sue@y.com" in recs               # co-recipient now harvested
        assert recs["sue@y.com"].num_recv == 0   # but counts 0 sent/received
        assert recs["sue@y.com"].num_sent == 0
        assert recs["sue@y.com"].names.get("Sue") == 1   # name captured
        assert "me@self.com" not in recs         # self never harvested

    def test_corecipient_who_is_sender_not_double_counted(self):
        recs = defaultdict(Rec)
        # sender also appears in To: should be credited received once, not twice
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")],
                             to=[("Bob", "bob@x.com"), ("Me", "me@self.com")]),
                        recs, {"me@self.com"})
        assert recs["bob@x.com"].num_recv == 1

    def test_no_cc_skips_corecipients(self):
        recs = defaultdict(Rec)
        _ingest_message(_msg(frm=[("Bob", "bob@x.com")],
                             to=[("Me", "me@self.com"), ("Sue", "sue@y.com")]),
                        recs, {"me@self.com"}, include_cc=False)
        assert recs["bob@x.com"].num_recv == 1   # sender still recorded
        assert "sue@y.com" not in recs           # co-recipient skipped


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


# LinkedIn export header + a 3-line "Notes:" preamble, like the real file.
_LI_PREAMBLE = ('Notes:\n'
                '"some privacy note about missing emails"\n'
                '\n'
                'First Name,Last Name,URL,Email Address,Company,Position,'
                'Connected On\n')


def _li_csv(rows):
    """rows: list of (first,last,url,email,company,position,connected)."""
    body = ""
    for r in rows:
        body += ",".join('"%s"' % c if "," in c else c for c in r) + "\n"
    return _LI_PREAMBLE + body


class TestParseLinkedin:
    def test_skips_preamble_and_parses(self):
        path = _write_tmp(_li_csv([
            ("Phil", "Dworsky", "https://www.linkedin.com/in/phildworsky", "",
             "GlobalFoundries", "Director, Strategic Programs", "05 Jun 2026"),
            ("Prashant", "Patil, PhD", "https://www.linkedin.com/in/drp", "",
             "Micromize, Inc", "CEO & Founder", "04 Jun 2026"),
        ]), ".csv")
        try:
            entries = parse_linkedin_csv(path)
        finally:
            os.remove(path)
        assert len(entries) == 2
        assert entries[0] == {
            "first": "Phil", "last": "Dworsky",
            "url": "https://www.linkedin.com/in/phildworsky", "email": "",
            "company": "GlobalFoundries", "position": "Director, Strategic Programs"}
        assert entries[1]["last"] == "Patil, PhD"      # quoted comma preserved
        assert entries[1]["company"] == "Micromize, Inc"

    def test_email_lowercased(self):
        path = _write_tmp(_li_csv([
            ("Sue", "Smith", "https://x/in/sue", "Sue@Y.COM", "Acme", "VP", "01 Jan 2020"),
        ]), ".csv")
        try:
            entries = parse_linkedin_csv(path)
        finally:
            os.remove(path)
        assert entries[0]["email"] == "sue@y.com"


class TestNameKey:
    def test_normalizes(self):
        assert _name_key("Jordan", "Vale") == "jordan vale"
        assert _name_key("Prashant", "Patil, PhD") == "prashant patil"   # suffix dropped
        assert _name_key(" JOHN ", "O'Brien") == "john o"
        assert _name_key("Cher", "") == ""           # need both names


class TestLoadRowsKeepsLinkedinOnly:
    def test_emailless_with_linkedin_kept(self):
        rows = [
            {"first_name": "No", "last_name": "Email", "primary_email": "",
             "emails": [], "linkedin": "https://x/in/noemail"},
            {"first_name": "Drop", "last_name": "Me", "primary_email": "",
             "emails": []},                       # no email, no linkedin -> dropped
        ]
        path = _write_tmp(json.dumps(rows), ".json")
        try:
            got = load_rows(path)
        finally:
            os.remove(path)
        assert len(got) == 1 and got[0]["last_name"] == "Email"


class TestFoldLinkedin:
    def _db(self):
        return [
            _row(first="Jordan", last="Vale", company="OldCo", title="",
                 primary="jordan@oldco.com", n_emails=5, last_i="2020-01-01"),
            _row(first="Sam", last="Lee", primary="sam1@a.com"),
            _row(first="Sam", last="Lee", primary="sam2@b.com"),   # name collision
        ]

    def test_match_by_name_overwrites_authority(self):
        rows = self._db()
        rows, enr, add, amb, skip = _fold_linkedin(rows, [
            {"first": "Jordan", "last": "Vale", "url": "https://x/in/jv",
             "email": "", "company": "Globex", "position": "VP Eng"}],
            "2026-06-06")
        j = next(r for r in rows if r["first_name"] == "Jordan")
        assert j["company"] == "Globex" and j["title"] == "VP Eng"  # overwritten
        assert j["linkedin"] == "https://x/in/jv"
        assert j["import_date"] == "2026-06-06"
        assert j["last_interaction"] == "2020-01-01"   # email recency preserved
        assert (enr, add, amb, skip) == (1, 0, 0, 0)

    def test_ambiguous_name_skipped(self):
        rows = self._db()
        before = len(rows)
        rows, enr, add, amb, skip = _fold_linkedin(rows, [
            {"first": "Sam", "last": "Lee", "url": "https://x/in/sam",
             "email": "", "company": "Acme", "position": "Eng"}], "2026-06-06")
        assert (enr, add, amb, skip) == (0, 0, 1, 0)
        assert len(rows) == before                     # not enriched, not added

    def test_unmatched_added_emailless(self):
        rows = self._db()
        rows, enr, add, amb, skip = _fold_linkedin(rows, [
            {"first": "Nora", "last": "New", "url": "https://x/in/nora",
             "email": "", "company": "Startup", "position": "Founder"}], "2026-06-06")
        n = next(r for r in rows if r["first_name"] == "Nora")
        assert add == 1 and n["primary_email"] == "" and n["emails"] == []
        assert n["linkedin"] == "https://x/in/nora" and n["company"] == "Startup"
        assert n["source"] == "linkedin"

    def test_no_email_no_url_skipped(self):
        rows = self._db()
        before = len(rows)
        rows, enr, add, amb, skip = _fold_linkedin(rows, [
            {"first": "Ghost", "last": "Person", "url": "", "email": "",
             "company": "Nowhere", "position": "Mystery"}], "2026-06-06")
        assert (enr, add, amb, skip) == (0, 0, 0, 1)   # no key -> skipped
        assert len(rows) == before

    def test_match_by_email_when_name_differs(self):
        rows = self._db()
        rows, enr, add, amb, skip = _fold_linkedin(rows, [
            {"first": "J", "last": "V", "url": "https://x/in/jv2",
             "email": "jordan@oldco.com", "company": "NewCo", "position": "CTO"}],
            "2026-06-06")
        j = next(r for r in rows if r["primary_email"] == "jordan@oldco.com")
        assert j["company"] == "NewCo" and add == 0 and enr == 1

    def test_idempotent_on_url(self):
        rows = self._db()
        entry = [{"first": "Nora", "last": "New", "url": "https://x/in/nora",
                  "email": "", "company": "Startup", "position": "Founder"}]
        rows, *_ = _fold_linkedin(rows, entry, "2026-06-06")
        n1 = len(rows)
        rows, enr, add, amb, skip = _fold_linkedin(rows, entry, "2026-06-07")
        assert add == 0 and len(rows) == n1            # url match, no duplicate
        n = next(r for r in rows if r["first_name"] == "Nora")
        assert n["import_date"] == "2026-06-07"        # refreshed


class TestLinkedinImportCli:
    def test_merge_into_db(self):
        db = [_row(first="Jordan", last="Vale", company="OldCo",
                   primary="jordan@oldco.com", n_emails=3, last_i="2020-01-01")]
        dbp = _write_tmp(json.dumps(db), ".json")
        lip = _write_tmp(_li_csv([
            ("Jordan", "Vale", "https://x/in/jv", "", "Globex", "VP", "05 Jun 2026"),
            ("Nora", "New", "https://x/in/nora", "", "Startup", "Founder", "01 Jan 2024"),
        ]), ".csv")
        try:
            main(["-i", lip, "--iformat", "linkedin", "-o", dbp,
                  "--import-date", "2026-06-06"])
            got = load_rows(dbp)
        finally:
            for p in (dbp, lip):
                os.remove(p)
        by_name = {(r["first_name"], r["last_name"]): r for r in got}
        assert len(got) == 2
        assert by_name[("Jordan", "Vale")]["company"] == "Globex"
        assert by_name[("Jordan", "Vale")]["import_date"] == "2026-06-06"
        assert by_name[("Nora", "New")]["primary_email"] == ""      # email-less new
        assert by_name[("Nora", "New")]["linkedin"] == "https://x/in/nora"


class TestMergeIsDefault:
    def _mbox(self):
        return ("From 1@xxx Wed Jun 03 14:35:08 +0000 2026\n"
                "Delivered-To: me@self.com\n"
                "From: Bob Jones <bob@x.com>\n"
                "To: me@self.com\n"
                "Date: Wed, 03 Jun 2026 14:35:06 +0000\n"
                'Content-Type: text/plain; charset="utf-8"\n'
                "\nhi\n")

    def test_import_does_not_wipe_existing_db(self):
        # Pre-existing DB with a hand-edited contact.
        db = [_row(first="Ann", last="Base", company="Acme", ctype="customer",
                   primary="ann@acme.com", n_recv=1)]
        dbp = _write_tmp(json.dumps(db), ".json")
        mp = _write_tmp(self._mbox(), ".mbox")
        try:
            main(["-i", mp, "-o", dbp])           # no --merge flag exists anymore
            got = {r["primary_email"]: r for r in load_rows(dbp)}
        finally:
            for p in (dbp, mp):
                os.remove(p)
        assert "ann@acme.com" in got              # existing kept (not wiped)
        assert got["ann@acme.com"]["type"] == "customer"
        assert "bob@x.com" in got                 # new import folded in

    def test_db_to_json_folds_not_overwrites(self):
        base = [_row(first="Ann", last="A", primary="ann@x.com", n_recv=1)]
        extra = [_row(first="Bob", last="B", primary="bob@y.com", n_recv=1)]
        basep = _write_tmp(json.dumps(base), ".json")
        extrap = _write_tmp(json.dumps(extra), ".json")
        try:
            main(["-i", extrap, "-o", basep])     # DB -> json merges into base
            emails = {r["primary_email"] for r in load_rows(basep)}
        finally:
            for p in (basep, extrap):
                os.remove(p)
        assert emails == {"ann@x.com", "bob@y.com"}


class TestNoCcCli:
    def _mbox(self):
        return ("From 1@xxx Wed Jun 03 14:35:08 +0000 2026\n"
                "Delivered-To: me@self.com\n"
                "From: Bob Jones <bob@x.com>\n"
                "To: me@self.com, Sue Smith <sue@y.com>\n"
                "Date: Wed, 03 Jun 2026 14:35:06 +0000\n"
                'Content-Type: text/plain; charset="utf-8"\n'
                "\nhi\n")

    def test_default_includes_cc_and_no_cc_excludes(self):
        mp = _write_tmp(self._mbox(), ".mbox")
        d1 = _write_tmp("", ".json")
        d2 = _write_tmp("", ".json")
        try:
            main(["-i", mp, "-o", d1])              # default: co-recipients in
            main(["-i", mp, "-o", d2, "--no-cc"])   # opt out
            with_cc = {r["primary_email"] for r in load_rows(d1)}
            without = {r["primary_email"] for r in load_rows(d2)}
        finally:
            for p in (mp, d1, d2):
                os.remove(p)
        assert with_cc == {"bob@x.com", "sue@y.com"}   # Sue (Cc) harvested
        assert without == {"bob@x.com"}                # Sue dropped with --no-cc


class TestImportDateStamp:
    def _mbox(self):
        return ("From 1@xxx Wed Jun 03 14:35:08 +0000 2026\n"
                "Delivered-To: me@self.com\n"
                "From: Bob Jones <bob@x.com>\n"
                "To: me@self.com\n"
                "Date: Wed, 03 Jun 2026 14:35:06 +0000\n"
                'Content-Type: text/plain; charset="utf-8"\n'
                "\nhi\n")

    def test_mbox_import_stamps_date(self):
        mp = _write_tmp(self._mbox(), ".mbox")
        outp = _write_tmp("", ".json")
        try:
            main(["-i", mp, "-o", outp, "--import-date", "2026-06-06"])
            got = load_rows(outp)
        finally:
            for p in (mp, outp):
                os.remove(p)
        assert got and all(r["import_date"] == "2026-06-06" for r in got)

    def test_db_convert_does_not_restamp(self):
        rows = [_row(first="A", last="One", primary="a@x.com", n_recv=1)]
        rows[0]["import_date"] = "2020-01-01"
        src = _write_tmp(json.dumps(rows), ".json")
        outp = _write_tmp("", ".json")
        try:
            main(["-i", src, "-o", outp])          # native json->json convert
            got = load_rows(outp)
        finally:
            for p in (src, outp):
                os.remove(p)
        assert got[0]["import_date"] == "2020-01-01"   # untouched


class TestReconcile:
    def test_merge_by_shared_email_diff_names(self):
        rows = [
            _row(first="Bob", last="Jones", primary="bob@acme.com",
                 emails=["bob@acme.com"], n_recv=2, last_i="2024-01-01"),
            _row(first="Robert", last="Jones", primary="rob@gmail.com",
                 emails=["rob@gmail.com", "bob@acme.com"], last_i="2025-06-01"),
        ]
        out = reconcile_contacts(rows)
        assert len(out) == 1
        r = out[0]
        assert "bob@acme.com" in r["emails"] and "rob@gmail.com" in r["emails"]
        assert (r["first_name"], r["last_name"]) == ("Robert", "Jones")  # newest

    def test_no_merge_on_shared_free_address(self):
        rows = [
            _row(first="Ann", last="One", primary="shared@gmail.com",
                 emails=["shared@gmail.com"]),
            _row(first="Bob", last="Two", primary="shared@gmail.com",
                 emails=["shared@gmail.com"]),
        ]
        out = reconcile_contacts(rows)
        assert len(out) == 2          # free provider not used as a merge key

    def test_linkedin_authority_company_title(self):
        a = _row(first="Bob", last="Jones", company="OldCo", title="Eng",
                 primary="bob@acme.com", emails=["bob@acme.com"], last_i="2025-01-01")
        b = _row(first="Bob", last="Jones", company="Acme Corp", title="VP",
                 primary="bob2@acme.com", emails=["bob2@acme.com", "bob@acme.com"],
                 last_i="2024-01-01")
        b["linkedin"] = "https://x/in/bob"
        out = reconcile_contacts([a, b])
        assert len(out) == 1
        assert out[0]["company"] == "Acme Corp" and out[0]["title"] == "VP"

    def test_primary_matches_company(self):
        r = _row(first="Sue", last="Smith", company="Nvidia",
                 primary="sue@gmail.com",
                 emails=["sue@gmail.com", "sue@nvidia.com"])
        out = reconcile_contacts([r])
        assert out[0]["primary_email"] == "sue@nvidia.com"

    def test_drops_junk_addresses(self):
        r = _row(first="Joe", last="Doe", primary="joe@acme.com",
                 emails=["joe@acme.com", "no-reply@acme.com", "bad-addr"])
        out = reconcile_contacts([r])
        assert out[0]["emails"] == ["joe@acme.com"]
        assert out[0]["primary_email"] == "joe@acme.com"

    def test_keeps_linkedin_only(self):
        r = _row(first="Lia", last="Only")
        r["primary_email"], r["emails"] = "", []
        r["linkedin"] = "https://x/in/lia"
        out = reconcile_contacts([r])
        assert len(out) == 1 and out[0]["linkedin"] == "https://x/in/lia"

    def test_drops_emailless_no_linkedin(self):
        r = _row(first="Ghost", last="None")
        r["primary_email"], r["emails"] = "", []
        assert reconcile_contacts([r]) == []

    def test_num_emails_recomputed(self):
        r = _row(first="Ed", last="Bee", primary="ed@x.com", emails=["ed@x.com"],
                 n_sent=2, n_recv=3, n_emails=999)
        out = reconcile_contacts([r])
        assert out[0]["num_emails"] == 5

    def test_name_titlecased(self):
        r = _row(first="BOB", last="jones", primary="b@x.com", emails=["b@x.com"])
        out = reconcile_contacts([r])
        assert (out[0]["first_name"], out[0]["last_name"]) == ("Bob", "Jones")

    def test_idempotent(self):
        rows = [
            _row(first="Bob", last="Jones", primary="bob@acme.com",
                 emails=["bob@acme.com"], n_recv=1, last_i="2024-01-01"),
            _row(first="Robert", last="Jones", primary="rob@gmail.com",
                 emails=["rob@gmail.com", "bob@acme.com"], last_i="2025-01-01"),
        ]
        once = reconcile_contacts(rows)
        twice = reconcile_contacts([dict(r) for r in once])
        assert len(once) == len(twice) == 1


class TestPipelineEndToEnd:
    # PST is omitted: libpff only reads real binary .pst files (no writer to build
    # a fixture); PST parsing is covered by TestPstMessageFields.
    def test_full_flow(self):
        mbox = ("From 1@xxx Wed Jun 03 14:35:08 +0000 2026\n"
                "Delivered-To: me@self.com\n"
                "From: Bob Jones <bob@acme.com>\n"
                "To: me@self.com\n"
                "Date: Wed, 03 Jun 2026 14:35:06 +0000\n"
                'Content-Type: text/plain; charset="utf-8"\n'
                "\nhi\n")
        vcf = ("BEGIN:VCARD\nVERSION:3.0\nFN:Robert Jones\nN:Jones;Robert;;;\n"
               "EMAIL;TYPE=INTERNET,PREF:rob.jones@gmail.com\n"
               "EMAIL;TYPE=INTERNET:bob@acme.com\nEND:VCARD\n")
        outlook = ("First Name,Last Name,Job Title,Company,E-mail Address\n"
                   "Carol,Lee,,Beta,carol@beta.com\n"
                   "Acme,Info,,Acme,info@acme.com\n")
        li = _li_csv([("Robert", "Jones", "https://x/in/robjones", "",
                       "Acme Corporation", "VP Sales", "01 Jan 2026")])
        mp = _write_tmp(mbox, ".mbox")
        vp = _write_tmp(vcf, ".vcf")
        op = _write_tmp(outlook, ".csv")
        lp = _write_tmp(li, ".csv")
        db = _write_tmp("", ".json")
        xl = _write_tmp("", ".xlsx")
        try:
            main(["-i", mp, "-o", db])                                  # mbox
            main(["-i", vp, "-o", db])                                  # vcard
            main(["-i", op, "--iformat", "outlook", "-o", db])         # outlook
            main(["-i", lp, "--iformat", "linkedin", "-o", db,         # linkedin
                  "--import-date", "2026-06-06"])
            main(["-i", db, "-o", db, "--reconcile"])                   # reconcile
            after = load_rows(db)
            main(["-i", db, "-o", xl])                                  # export xlsx
            xl_rows = load_rows(xl)
        finally:
            for p in (mp, vp, op, lp, db, xl):
                os.remove(p)
        emails = {r["primary_email"] for r in after}
        # mbox 'Bob Jones' + vCard 'Robert Jones' share bob@acme.com -> 1 record
        jones = [r for r in after if r["last_name"] == "Jones"]
        assert len(jones) == 1
        j = jones[0]
        assert "bob@acme.com" in j["emails"] and "rob.jones@gmail.com" in j["emails"]
        assert j["company"] == "Acme Corporation"          # LinkedIn authority
        assert j["linkedin"] == "https://x/in/robjones"
        # Carol Lee flows through; the role address info@acme.com is dropped
        assert "carol@beta.com" in emails
        assert not any(e == "info@acme.com" for r in after for e in r["emails"])
        assert not any(r["last_name"] == "Info" for r in after)
        # xlsx export round-trips equal to the reconciled json
        assert xl_rows == after
