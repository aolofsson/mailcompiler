#!/usr/bin/env python3
"""mailcompiler (mc): build and query a contacts database.

Single CLI with one function per operation, dispatched as a subcommand:
  mc import  -i MBOX|PST -o OUT [...]           build/merge the contacts DB
  mc list    -i CONTACTS [filters...]           list matching addresses

Import accepts a Gmail Takeout .mbox or an Outlook .pst (chosen by extension).
The database is JSON (the native format); pass -o something.csv to export CSV.

mbox import streams the file (only header blocks are parsed, bodies skipped),
so a 21 GB mbox is handled in one pass with minimal memory. PST import uses
pypff (libpff-python).

Spec (confirmed with user):
  - Contacts = recipients (To/Cc) of mail I SENT, plus senders (From) of mail I
    RECEIVED. Self addresses are excluded.
  - Spam-labeled messages are skipped (unsolicited, not correspondence).
  - Automated/bulk senders (no-reply, mailer-daemon, bulk ESP domains, ...) are
    filtered out.
  - Identities merged by display name (multiple addresses -> one person).
  - Company derived from email domain, blank for free providers.
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import defaultdict
from datetime import timezone
from email.parser import HeaderParser
from email.header import decode_header, make_header
from email.utils import getaddresses, parsedate_to_datetime

SEP = re.compile(r'^From \d+@xxx ')

FREE_PROVIDERS = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.co.uk", "ymail.com",
    "outlook.com", "hotmail.com", "live.com", "msn.com", "icloud.com",
    "me.com", "mac.com", "aol.com", "proton.me", "protonmail.com", "gmx.com",
    "gmx.net", "mail.com", "zoho.com", "yandex.com", "fastmail.com",
    "comcast.net", "verizon.net", "sbcglobal.net", "att.net", "qq.com",
    "163.com", "126.com",
}

# Local-part patterns that indicate an automated / non-human sender. Matched
# against a normalized local-part (._ collapsed to -).
BOT_LOCAL = re.compile(
    r'^(no-?reply|do-?not-?reply|donotreply|mailer-daemon|mailer|postmaster|'
    r'bounce|bounces|notification|notifications|notify|noreply|alert|alerts|'
    r'newsletter|newsletters|mailing|mailings|nepliy|automated|'
    r'auto-?reply|autoreply|root|daemon|cron)$'
)
BOT_SUBSTR = re.compile(
    r'(no-?reply|noreply|do-?not-?reply|donotreply|mailer-daemon|'
    r'marketing|newsletter|mailings?|unsub|spamproc)')
# Local-part prefixes (before a "+" tag) that indicate automated mail,
# e.g. GitHub's reply+<hash>@reply.github.com or notifications+<id>@...
BOT_TAG_PREFIX = {
    "reply", "notify", "notification", "notifications", "bounce", "bounces",
    "comment", "comments", "mention", "mentions", "noreply", "no-reply",
}

# Bulk email service provider / list domains (suffix or prefix match).
BOT_DOMAINS = (
    "mailchimp.com", "mcsv.net", "mcdlv.net", "rsgsv.net", "list-manage.com",
    "sendgrid.net", "sendgrid.com", "mailgun.org", "mailgun.com", "sparkpostmail.com",
    "amazonses.com", "mandrillapp.com", "sendinblue.com", "sib.email",
    "substack.com", "bounce.com", "bounces.google.com", "mailer.com",
    "constantcontact.com", "ccsend.com", "hubspotemail.net", "hs-send.com",
    "salesforce.com", "exct.net", "rs6.net", "klaviyomail.com", "intercom-mail.com",
    "mixmax.com", "mail.notion.so", "reply.github.com", "github.com",
    "mktomail.com", "beehiiv.com", "en25.com", "hubspotstarter.net",
    "connectedcommunity.org", "mailmarketo.com", "marketo.com",
    "notifications.", "noreply.", "engagement.", "marketing.", "email.",
    "news.", "mailing.", "updates.", "em.",
)


def dec(s):
    """Decode an RFC2047 header value to plain text."""
    if not s:
        return ""
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return s


def is_bot(email_addr):
    local, _, domain = email_addr.partition("@")
    if not domain:
        return True
    norm = local.lower().replace("_", "-").replace(".", "-")
    if BOT_LOCAL.match(norm) or BOT_SUBSTR.search(local.lower()):
        return True
    if "+" in local and local.split("+", 1)[0].lower() in BOT_TAG_PREFIX:
        return True
    for d in BOT_DOMAINS:
        if domain == d or domain.endswith("." + d) or domain.startswith(d):
            return True
    return False


def load_blacklist(path):
    """Read a blacklist file into a set of domains (one entry per line).

    Blank lines and lines starting with '#' are ignored. Entries may be written
    as "example.com" or "@example.com"; both mean "block this whole domain".
    """
    domains = set()
    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            entry = line.strip()
            if not entry or entry.startswith("#"):
                continue
            domains.add(entry.lstrip("@").lower())
    return domains


def is_blacklisted(email_addr, blacklist_domains):
    """True if the address's domain matches a blacklisted domain or subdomain."""
    if not blacklist_domains:
        return False
    domain = email_addr.split("@")[-1].lower()
    return any(domain == d or domain.endswith("." + d) for d in blacklist_domains)


def clean(tok):
    """Strip stray wrapping quotes/punctuation from a name token."""
    return tok.strip(" '\"().,-")


def split_name(display, email_addr):
    """Return (first, last) from a display name, with email fallback."""
    name = (display or "").strip().strip('"\'').strip()
    # Cut org/department/handle suffixes leaked into the display name, e.g.
    # "Seo (Ethan)/dept/SUPEX" or "Name <tag>" -> keep the leading human part.
    name = re.split(r'\s*[(/\\<|]', name)[0].strip().rstrip(").-").strip()
    # No usable display name: synthesize from the email local-part, but only
    # from human-looking tokens (alpha, optional trailing digits, <=20 chars).
    # This recovers "eric.wallace.4" -> Eric Wallace while rejecting hash/unsub
    # local-parts like "8a795fa6-038a-4166-...".
    if not name or ("@" in name and " " not in name):
        local = email_addr.split("@")[0]
        toks = [t for t in re.split(r'[._\-+]+', local)
                if re.fullmatch(r'[A-Za-z]{2,}\d*', t) and len(t) <= 20]
        if len(toks) >= 2:
            return toks[0].capitalize(), toks[-1].capitalize()
        if len(toks) == 1:
            return toks[0].capitalize(), ""
        return "", ""
    if "," in name:  # "Last, First"
        last, _, first = name.partition(",")
        ftok = clean(first).split()
        return (ftok[0].capitalize() if ftok else ""), clean(last).capitalize()
    parts = clean(name).split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[-1]


def company_from(email_addr):
    domain = email_addr.split("@")[-1].lower()
    if domain in FREE_PROVIDERS:
        return ""
    base = domain.split(".")
    if len(base) >= 2:
        return base[-2].capitalize()
    return domain.capitalize()


# ---- per-address accumulator ------------------------------------------------
class Rec:
    __slots__ = ("emails", "names", "num_sent", "num_recv", "first", "last")

    def __init__(self):
        self.emails = defaultdict(int)   # email -> count
        self.names = defaultdict(int)    # display name -> count
        self.num_sent = 0                # I -> them
        self.num_recv = 0                # them -> I
        self.first = None
        self.last = None

    def touch(self, dt):
        if dt is None:
            return
        if self.first is None or dt < self.first:
            self.first = dt
        if self.last is None or dt > self.last:
            self.last = dt


# ---- CSV output / additive merge -------------------------------------------
CSV_FIELDS = ["vip", "last_name", "first_name", "company", "primary_email",
              "emails", "num_emails", "num_sent", "num_received",
              "first_interaction", "last_interaction", "source"]

# Source values may contain spaces (mbox filenames), so they are joined with
# this separator rather than a space (which the emails column uses).
SOURCE_SEP = " | "


def _to_int(v):
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return 0


def person_to_row(p, source):
    """Convert an in-memory person dict to a CSV row dict, tagged with source."""
    return {
        "vip": "",
        "last_name": p["last_name"],
        "first_name": p["first_name"],
        "company": p["company"],
        "primary_email": p["primary_email"],
        "emails": list(p["emails"]),
        "num_emails": p["num_emails"],
        "num_sent": p["num_sent"],
        "num_received": p["num_received"],
        "first_interaction": p["first_interaction"],
        "last_interaction": p["last_interaction"],
        "source": source,
    }


def _normalize_row(d):
    """Coerce a loaded record (from JSON or CSV) into the canonical row dict.

    Accepts `emails` as either a list (JSON) or a space-separated string (CSV).
    """
    emails = d.get("emails")
    if not isinstance(emails, list):
        emails = (emails or "").split()
    return {
        "vip": str(d.get("vip") or "").strip(),
        "last_name": str(d.get("last_name") or "").strip(),
        "first_name": str(d.get("first_name") or "").strip(),
        "company": str(d.get("company") or "").strip(),
        "primary_email": str(d.get("primary_email") or "").strip(),
        "emails": list(emails),
        "num_emails": _to_int(d.get("num_emails")),
        "num_sent": _to_int(d.get("num_sent")),
        "num_received": _to_int(d.get("num_received")),
        "first_interaction": (str(d.get("first_interaction") or "").strip() or None),
        "last_interaction": (str(d.get("last_interaction") or "").strip() or None),
        "source": str(d.get("source") or "").strip(),
    }


def load_rows(path):
    """Load a contacts file into a list of canonical row dicts.

    Format follows the extension: .json (the native store) or .csv. Missing
    files yield an empty list.
    """
    if not os.path.isfile(path):
        return []
    if path.lower().endswith(".json"):
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        records = data if isinstance(data, list) else []
    else:
        with open(path, "r", encoding="utf-8", newline="") as fh:
            records = list(csv.DictReader(fh))
    return [_normalize_row(d) for d in records
            if str(d.get("primary_email") or "").strip()]


def read_existing_rows(path):
    """Read an existing contacts file into {primary_email_lower: row dict}."""
    return {r["primary_email"].lower(): r for r in load_rows(path)}


def _merge_date(a, b, newest):
    vals = [d for d in (a, b) if d]
    if not vals:
        return None
    return max(vals) if newest else min(vals)


def _union_sources(a, b):
    """Union two SOURCE_SEP-joined source strings, de-duped, order-preserving."""
    out = []
    for s in a.split(SOURCE_SEP) + b.split(SOURCE_SEP):
        s = s.strip()
        if s and s not in out:
            out.append(s)
    return SOURCE_SEP.join(out)


def merge_row(existing, new):
    """Merge `new` into `existing`. Hand-edited text fields (vip, name, company)
    are preserved; counts are overwritten with the latest import; emails and
    sources union; the date range widens."""
    existing["vip"] = existing["vip"] or new["vip"]
    for f in ("last_name", "first_name", "company"):
        existing[f] = existing[f] or new[f]
    for e in new["emails"]:
        if e not in existing["emails"]:
            existing["emails"].append(e)
    existing["num_emails"] = new["num_emails"]
    existing["num_sent"] = new["num_sent"]
    existing["num_received"] = new["num_received"]
    existing["first_interaction"] = _merge_date(
        existing["first_interaction"], new["first_interaction"], newest=False)
    existing["last_interaction"] = _merge_date(
        existing["last_interaction"], new["last_interaction"], newest=True)
    existing["source"] = _union_sources(existing["source"], new["source"])


def _write_atomic(path, write_fn):
    """Run write_fn(file) against a temp file then os.replace, so a crash
    mid-write cannot corrupt an existing contacts file we just merged into."""
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as out:
        write_fn(out)
    os.replace(tmp, path)


def write_csv_rows(path, rows):
    def _w(out):
        w = csv.writer(out)
        w.writerow(CSV_FIELDS)
        for r in rows:
            w.writerow([
                r["vip"], r["last_name"], r["first_name"], r["company"],
                r["primary_email"], " ".join(r["emails"]), r["num_emails"],
                r["num_sent"], r["num_received"],
                r["first_interaction"] or "", r["last_interaction"] or "",
                r.get("source", ""),
            ])
    _write_atomic(path, _w)


def write_json_rows(path, rows):
    ordered = [{f: r.get(f) for f in CSV_FIELDS} for r in rows]

    def _w(out):
        json.dump(ordered, out, indent=2, ensure_ascii=False)
        out.write("\n")
    _write_atomic(path, _w)


def write_rows(path, rows):
    """Write contacts to `path`. JSON is the native store; a .csv path exports
    the same records as CSV."""
    if path.lower().endswith(".csv"):
        write_csv_rows(path, rows)
    else:
        write_json_rows(path, rows)


# ---- query / list helpers ---------------------------------------------------
def _csv_set(value):
    """Parse a comma-separated CLI value into a lowercased set, or None."""
    if value is None:
        return None
    return {v.strip().lower() for v in value.split(",") if v.strip()}


def _domain(addr):
    return addr.split("@")[-1].lower() if addr else ""


def build_criteria(args):
    """Turn parsed list args into a flat dict of active predicates."""
    return {
        "vip": args.vip,
        "company": _csv_set(args.company),
        "first_name": _csv_set(args.first_name),
        "last_name": _csv_set(args.last_name),
        "email_domain": _csv_set(args.email_domain),
        "min_emails": args.min_emails,
        "max_emails": args.max_emails,
        "min_sent": args.min_sent,
        "max_sent": args.max_sent,
        "min_received": args.min_received,
        "max_received": args.max_received,
        "last_after": args.last_after,
        "last_before": args.last_before,
        "first_after": args.first_after,
        "first_before": args.first_before,
    }


def matches(contact, crit):
    """Return True if a contact satisfies every active criterion (AND)."""
    # VIP flag: only contacts with a non-empty vip column.
    if crit["vip"] and not str(contact.get("vip", "")).strip():
        return False

    # Text set-membership (case-insensitive exact).
    for field in ("company", "first_name", "last_name"):
        allowed = crit[field]
        if allowed is not None and str(contact.get(field, "")).lower() not in allowed:
            return False
    if crit["email_domain"] is not None:
        if _domain(contact.get("primary_email", "")) not in crit["email_domain"]:
            return False

    # Numeric inclusive ranges.
    for field, lo, hi in (
        ("num_emails", crit["min_emails"], crit["max_emails"]),
        ("num_sent", crit["min_sent"], crit["max_sent"]),
        ("num_received", crit["min_received"], crit["max_received"]),
    ):
        val = contact.get(field, 0) or 0
        if lo is not None and val < lo:
            return False
        if hi is not None and val > hi:
            return False

    # Date inclusive ranges (ISO YYYY-MM-DD strings; null date fails any bound).
    for field, after, before in (
        ("last_interaction", crit["last_after"], crit["last_before"]),
        ("first_interaction", crit["first_after"], crit["first_before"]),
    ):
        if after is None and before is None:
            continue
        val = contact.get(field)
        if not val:
            return False
        if after is not None and val < after:
            return False
        if before is not None and val > before:
            return False

    return True


def select_addresses(contacts, all_emails):
    """Collect primary (or all) addresses, de-duplicated in first-seen order."""
    seen = set()
    out = []
    for c in contacts:
        addrs = c.get("emails", []) if all_emails else [c.get("primary_email", "")]
        for a in addrs:
            a = (a or "").strip()
            if a and a.lower() not in seen:
                seen.add(a.lower())
                out.append(a)
    return out


# ---- message ingestion ------------------------------------------------------
# Each source (mbox / PST) yields a normalized message dict:
#   {"from": [(name, addr), ...], "to": [(name, addr), ...]  (To+Cc),
#    "date": <naive-UTC datetime or None>, "is_sent": bool, "is_spam": bool,
#    "self_hints": [addr, ...]}  -- addresses known to be the account owner.

# Outlook folder names that mark sent / spam mail (case-insensitive).
PST_SENT_FOLDERS = {"sent items", "sent", "sent mail"}
PST_SPAM_FOLDERS = {"junk email", "junk e-mail", "junk", "spam"}

# MAPI property tags used by the PST fallback when there are no RFC822 headers.
_MAPI_SENDER_NAME = 0x0C1A
_MAPI_SENDER_SMTP = 0x5D01
_MAPI_SENDER_EMAIL = 0x0C1F


def _normalize_dt(dt):
    """Coerce a datetime to naive UTC (or return None)."""
    if dt is None:
        return None
    try:
        if dt.tzinfo is not None:
            dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    except Exception:
        return None
    return dt


def _ingest_message(msg, recs, self_set):
    """Fold one normalized message into `recs`/`self_set`.

    Returns False if the message was skipped (spam), True otherwise.
    """
    if msg["is_spam"]:
        return False
    for a in msg["self_hints"]:
        if a:
            self_set.add(a.lower())

    from_emails = [addr.lower() for _, addr in msg["from"] if addr]
    sent_by_me = msg["is_sent"] or any(a in self_set for a in from_emails)
    dt = msg["date"]
    if sent_by_me:
        for a in from_emails:        # learn self addresses
            self_set.add(a)
        pairs, attr = msg["to"], "num_sent"
    else:
        pairs, attr = msg["from"], "num_recv"

    for raw_name, addr in pairs:
        addr = (addr or "").lower().strip()
        if not addr or "@" not in addr or addr in self_set:
            continue
        r = recs[addr]
        r.emails[addr] += 1
        nm = dec(raw_name).strip()
        if nm and "@" not in nm:
            r.names[nm] += 1
        setattr(r, attr, getattr(r, attr) + 1)
        r.touch(dt)
    return True


def iter_mbox_messages(path):
    """Yield normalized messages from a Gmail Takeout mbox (streamed)."""
    parser = HeaderParser()

    def normalize(block):
        m = parser.parsestr("".join(block))
        labels = {x.strip() for x in (m.get("X-Gmail-Labels") or "").split(",")}
        try:
            dt = _normalize_dt(parsedate_to_datetime(m.get("Date")))
        except Exception:
            dt = None
        return {
            "from": getaddresses(m.get_all("From", [])),
            "to": getaddresses(m.get_all("To", []) + m.get_all("Cc", [])),
            "date": dt,
            "is_sent": "Sent" in labels,
            "is_spam": "Spam" in labels,
            "self_hints": [a for _, a in getaddresses(m.get_all("Delivered-To", [])) if a],
        }

    with open(path, "r", encoding="latin-1", errors="replace") as fh:
        buf = []
        in_header = True
        for line in fh:
            if SEP.match(line):
                if buf:
                    yield normalize(buf)
                buf = []
                in_header = True
                continue
            if in_header:
                if line in ("\n", "\r\n"):
                    in_header = False  # end of headers; ignore body
                else:
                    buf.append(line)
        if buf:
            yield normalize(buf)


def _pst_record_value(message, entry_type):
    """Return the string value of a MAPI property on a pypff message, or ''."""
    try:
        nsets = message.number_of_record_sets
    except Exception:
        return ""
    for i in range(nsets):
        try:
            rs = message.get_record_set(i)
            for j in range(rs.number_of_entries):
                entry = rs.get_entry(j)
                if entry.entry_type == entry_type:
                    return (entry.get_data_as_string() or "").strip()
        except Exception:
            continue
    return ""


def _pst_date(message):
    for attr in ("client_submit_time", "delivery_time"):
        try:
            dt = getattr(message, attr)
        except Exception:
            dt = None
        if dt is not None:
            return _normalize_dt(dt)
    return None


def _pst_message_fields(message, folder_name):
    """Normalize one pypff message (in `folder_name`) into a message dict.

    Prefers the original RFC822 transport headers (present for most mail) and
    parses them like an mbox message; otherwise falls back to MAPI properties
    for the sender. Recipients are only recovered from transport headers, so
    Outlook-composed mail without headers contributes its sender only.
    """
    fname = (folder_name or "").strip().lower()
    is_sent = fname in PST_SENT_FOLDERS
    is_spam = fname in PST_SPAM_FOLDERS

    try:
        headers = message.transport_headers
    except Exception:
        headers = None

    if headers:
        m = HeaderParser().parsestr(headers)
        try:
            dt = _normalize_dt(parsedate_to_datetime(m.get("Date")))
        except Exception:
            dt = None
        return {
            "from": getaddresses(m.get_all("From", [])),
            "to": getaddresses(m.get_all("To", []) + m.get_all("Cc", [])),
            "date": dt,
            "is_sent": is_sent,
            "is_spam": is_spam,
            "self_hints": [],
        }

    # MAPI fallback (no transport headers).
    name = getattr(message, "sender_name", None) or \
        _pst_record_value(message, _MAPI_SENDER_NAME)
    email = (_pst_record_value(message, _MAPI_SENDER_SMTP)
             or _pst_record_value(message, _MAPI_SENDER_EMAIL))
    if "@" not in email:
        email = ""
    return {
        "from": [(name or "", email)],
        "to": [],
        "date": _pst_date(message),
        "is_sent": is_sent,
        "is_spam": is_spam,
        "self_hints": [],
    }


def _walk_pst_folder(folder):
    """Recursively yield normalized messages from a pypff folder."""
    try:
        name = folder.name or ""
    except Exception:
        name = ""
    try:
        n_msgs = folder.number_of_sub_messages
    except Exception:
        n_msgs = 0
    for i in range(n_msgs):
        try:
            message = folder.get_sub_message(i)
        except Exception:
            continue
        yield _pst_message_fields(message, name)
    try:
        n_sub = folder.number_of_sub_folders
    except Exception:
        n_sub = 0
    for i in range(n_sub):
        try:
            sub = folder.get_sub_folder(i)
        except Exception:
            continue
        yield from _walk_pst_folder(sub)


def iter_pst_messages(path):
    """Yield normalized messages from an Outlook PST (requires pypff)."""
    try:
        import pypff
    except ImportError:
        sys.exit("error: reading .pst requires the 'libpff-python' package "
                 "(pip install libpff-python)")
    pst = pypff.file()
    pst.open(path)
    try:
        yield from _walk_pst_folder(pst.get_root_folder())
    finally:
        pst.close()


# ---- commands ---------------------------------------------------------------
def cmd_import(args):
    """Import a Gmail mbox or Outlook PST (args.input) into the contacts DB."""
    path = args.input
    # Self addresses are auto-detected: the mbox Delivered-To header / the From
    # of Sent-folder (or Sent-labeled) mail.
    self_set = set()

    if not os.path.isfile(path):
        sys.exit("error: input not found: %s" % path)

    blacklist = set()
    if args.blacklist:
        if not os.path.isfile(args.blacklist):
            sys.exit("error: blacklist not found: %s" % args.blacklist)
        blacklist = load_blacklist(args.blacklist)

    recs = defaultdict(Rec)  # primary email -> Rec
    n_msgs = 0
    n_skip_spam = 0

    # Pick the reader by extension: .pst -> Outlook, otherwise mbox.
    src = (iter_pst_messages(path) if path.lower().endswith(".pst")
           else iter_mbox_messages(path))
    for msg in src:
        n_msgs += 1
        if not _ingest_message(msg, recs, self_set):
            n_skip_spam += 1
        if n_msgs % 50000 == 0:
            sys.stderr.write(f"  parsed {n_msgs:,} messages\n")
            sys.stderr.flush()

    sys.stderr.write(f"Done parsing: {n_msgs:,} messages, {n_skip_spam:,} spam skipped.\n")
    sys.stderr.write(f"Self addresses: {sorted(self_set)}\n")

    # ---- drop self + blacklist + bots ---------------------------------------
    for a in list(recs):
        if a in self_set:
            del recs[a]
    blacklisted = {a for a in recs if is_blacklisted(a, blacklist)}
    for a in blacklisted:
        del recs[a]
    if blacklist:
        sys.stderr.write(
            f"Filtered {len(blacklisted):,} blacklisted addresses "
            f"({len(blacklist):,} domains).\n")
    bots = {a for a in recs if is_bot(a)}
    for a in bots:
        del recs[a]
    sys.stderr.write(f"Filtered {len(bots):,} automated/bulk addresses.\n")

    # ---- merge by display name ---------------------------------------------
    def canon_name(r):
        if not r.names:
            return None
        nm = max(r.names, key=r.names.get)
        return re.sub(r'\s+', ' ', nm.lower().replace(",", " ")).strip()

    by_name = defaultdict(list)
    no_name = []
    for addr, r in recs.items():
        key = canon_name(r)
        if key:
            by_name[key].append(addr)
        else:
            no_name.append(addr)

    people = []

    def build(addrs):
        merged = Rec()
        for a in addrs:
            r = recs[a]
            for e, c in r.emails.items():
                merged.emails[e] += c
            for nm, c in r.names.items():
                merged.names[nm] += c
            merged.num_sent += r.num_sent
            merged.num_recv += r.num_recv
            merged.touch(r.first)
            merged.touch(r.last)
        # primary email = most-used address
        primary = max(merged.emails, key=merged.emails.get)
        display = max(merged.names, key=merged.names.get) if merged.names else ""
        first, last = split_name(display, primary)
        return {
            "first_name": first,
            "last_name": last,
            "company": company_from(primary),
            "primary_email": primary,
            "emails": sorted(merged.emails, key=lambda e: -merged.emails[e]),
            "display_name": display,
            "num_emails": merged.num_sent + merged.num_recv,
            "num_sent": merged.num_sent,
            "num_received": merged.num_recv,
            "first_interaction": merged.first.date().isoformat() if merged.first else None,
            "last_interaction": merged.last.date().isoformat() if merged.last else None,
        }

    for key, addrs in by_name.items():
        people.append(build(addrs))
    for addr in no_name:
        people.append(build([addr]))

    # ---- final contact filters: full name + corresponded in either direction
    # Keep anyone you sent to OR received from. (Every recorded contact already
    # has at least one sent or received message, so the guard documents the
    # rule rather than dropping anyone today.)
    n_built = len(people)
    people = [p for p in people if p["first_name"] and p["last_name"]]
    n_after_name = len(people)
    people = [p for p in people if p["num_sent"] > 0 or p["num_received"] > 0]
    n_kept = len(people)
    sys.stderr.write(
        f"Contacts: {n_built:,} built -> {n_after_name:,} with full name "
        f"-> {n_kept:,} with correspondence.\n")

    # -o may be a directory (writes contacts.json) or a file path. The format
    # follows the extension: a .csv path exports CSV, otherwise JSON (native).
    out = os.path.abspath(args.out)
    if os.path.isdir(out) or args.out.endswith(os.sep):
        cpath = os.path.join(out, "contacts.json")
    else:
        cpath = out
        if os.path.splitext(cpath)[1].lower() not in (".json", ".csv"):
            cpath += ".json"
    outdir = os.path.dirname(cpath)
    if outdir:
        os.makedirs(outdir, exist_ok=True)

    # Merge into any existing file: counts are overwritten with this import,
    # emails union and the date range widens, and hand-edited fields (e.g. vip)
    # are preserved. Contacts present only in the old file are kept untouched.
    # With -f/--force, ignore the existing file and overwrite it.
    merged = {} if args.force else read_existing_rows(cpath)
    n_existing = len(merged)
    n_new = n_updated = 0
    source = os.path.basename(path)
    for p in people:
        row = person_to_row(p, source)
        key = row["primary_email"].lower()
        if key in merged:
            merge_row(merged[key], row)
            n_updated += 1
        else:
            merged[key] = row
            n_new += 1

    rows = sorted(merged.values(), key=lambda r: (
        r["company"] == "",
        r["company"].lower(),
        -r["num_emails"],
        r["last_name"].lower(),
        r["first_name"].lower(),
    ))
    write_rows(cpath, rows)

    verb = "Overwrote" if args.force else "Merged into"
    sys.stderr.write(
        f"\n{verb} {cpath}\n"
        f"  {n_existing:,} existing + {n_new:,} new "
        f"({n_updated:,} updated) = {len(rows):,} contacts.\n")


def cmd_list(args):
    """List matching addresses from a contacts file (JSON or CSV, args.input)."""
    contacts = load_rows(args.input)
    crit = build_criteria(args)
    selected = [c for c in contacts if matches(c, crit)]
    addrs = select_addresses(selected, args.all_emails)

    line = ", ".join(addrs)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as out:
            out.write(line + "\n")
    else:
        print(line)

    sys.stderr.write(
        f"Matched {len(selected):,}/{len(contacts):,} contacts, "
        f"{len(addrs):,} addresses.\n")


def parse_args(argv=None):
    p = argparse.ArgumentParser(
        prog="mc",
        description="mailcompiler: build and query a contacts database.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    # mc import -i MBOX -o OUT [...]
    sc = sub.add_parser(
        "import", help="import a Gmail mbox or Outlook PST into the database")
    sc.add_argument("-i", "--input", dest="input", required=True,
                    help="path to the .mbox or .pst file to import")
    sc.add_argument("-o", dest="out", required=True,
                    help="output path: .json (native) or .csv (export); a "
                         "directory writes contacts.json")
    sc.add_argument("--blacklist", dest="blacklist", metavar="PATH",
                    help="file of domains to exclude from contacts (one per "
                         "line; '#' comments and blank lines ignored)")
    sc.add_argument("-f", "--force", action="store_true",
                    help="overwrite the output file instead of merging into it")
    sc.set_defaults(func=cmd_import)

    # mc list -i CONTACTS [filters...]
    ex = sub.add_parser(
        "list", help="list matching addresses from the contacts database")
    ex.add_argument("-i", "--input", dest="input", required=True,
                    help="path to the contacts JSON or CSV from 'mc import'")
    ex.add_argument("-o", "--output", dest="output",
                    help="write the address line to a file (default: stdout)")
    ex.add_argument("--all-emails", action="store_true",
                    help="emit every address per contact (default: primary only)")
    ex.add_argument("--vip", action="store_true",
                    help="only contacts flagged VIP (non-empty vip column)")
    ex.add_argument("--company", help="match company against any of LIST")
    ex.add_argument("--first-name", dest="first_name",
                    help="match first name against any of LIST")
    ex.add_argument("--last-name", dest="last_name",
                    help="match last name against any of LIST")
    ex.add_argument("--email-domain", dest="email_domain",
                    help="match primary email domain against any of LIST")
    ex.add_argument("--min-emails", type=int, help="minimum num_emails")
    ex.add_argument("--max-emails", type=int, help="maximum num_emails")
    ex.add_argument("--min-sent", type=int, help="minimum num_sent")
    ex.add_argument("--max-sent", type=int, help="maximum num_sent")
    ex.add_argument("--min-received", type=int, help="minimum num_received")
    ex.add_argument("--max-received", type=int, help="maximum num_received")
    ex.add_argument("--last-after", dest="last_after", metavar="YYYY-MM-DD",
                    help="last_interaction on or after this date")
    ex.add_argument("--last-before", dest="last_before", metavar="YYYY-MM-DD",
                    help="last_interaction on or before this date")
    ex.add_argument("--first-after", dest="first_after", metavar="YYYY-MM-DD",
                    help="first_interaction on or after this date")
    ex.add_argument("--first-before", dest="first_before", metavar="YYYY-MM-DD",
                    help="first_interaction on or before this date")
    ex.set_defaults(func=cmd_list)

    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    main()
