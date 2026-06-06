![MailCompiler](docs/postverket.svg)

Turn your email archive into a contact list you own. MailCompiler imports your
`.mbox`/`.pst` into a clean, human-readable JSON database and converts between
vCard and CSV.

## Motivation
- It's getting increasinly more important to own and protect your data!
- Manually scraping inboxes to put together outreach lists is a waste of life.
- Email client search functiosn are completely useless.
- Email client import/output is are different shades of broken.

## Key features
- Clean human readable JSON contact database
- Fast import from Gmail Takeout `.mbox` and Outlook `.pst` (handles 20 GB+)
- Import/export support for VCARD (3.0) and CSV (Outlook) contact lists
- Lossless import/export between xls/csv files and JSON database
- Streaming .mbox into LLM friendly JSON corpus
- Automatic extraction of email conversations into contacts
- Automatic extraction of phone numbers from email signatures
- Incremental non-destructive merging of into a JSON database
- Deduplication of records
- Black list email list support
- Automatic filtering of bot farm email addresses
- Record filtering on export

## TL;DR

```bash
pip install -e .                                   # install the `mc` command
mc -i takeout.mbox -o contacts.json                # mailbox -> contacts DB
mc -i contacts.json -o contacts.xlsx               # edit in Excel, then...
mc -i contacts.xlsx -o contacts.json               # ...lossless round-trip back
mc -i contacts.json --type customer -o leads.vcf   # filter + export (csv/xlsx/vcf)
```

There's one command, `mc`, and one JSON database. `mc` infers the operation from
the `-i`/`-o` file extensions: a mailbox/vCard/CSV/XLSX/LinkedIn in **imports**;
a JSON DB in **exports** (or `--dedup`). Imports always fold into the existing
`-o` (never wiping it). The rest of this README is the details.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Workflow

1. Convert email archive (.mbox, .pst) into a JSON file
  mc -i "All mail Including Spam and Trash.mbox" -o contacts.json
2. Translate JSON file to an xlsx spreadsheet
  mc -i contacts.json -o contacts.xlsx
3. Edit/annotate/wrangle/fix contact spreadsheet
4. Merge changes into golden contact database
  mc -i contacts.xlsx -o contacts.json

## Examples

Build the JSON contacts database from a Gmail Takeout mbox:

    mc -i "All mail Including Spam and Trash.mbox" -o data/contacts.json

Also harvest the other To/Cc recipients of mail you received (people on a
thread with you, not just the sender), counted as `num_cc`:

    mc -i "All mail Including Spam and Trash.mbox" -o data/contacts.json --include-cc

Direct extraction from Takeout mbox to an xlsx spreadsheet:

    mc -i "All mail Including Spam and Trash.mbox" -o data/contacts.xlsx

Import an Outlook PST instead:

    mc -i archive.pst -o data/contacts.json

Convert a JSON contacts database to an excel spreadsheet

    mc -i data/contacts.json -o contacts.xlsx

Import a vCard export (Google Contacts / Gmail):

    mc -i contacts.vcf -o data/contacts.json

Import an Outlook / Google Contacts CSV export (`--iformat outlook`):

    mc -i contacts.csv --iformat outlook -o data/contacts.json

Enrich the DB from a LinkedIn Connections export (`--iformat linkedin`):
overwrites company/title (LinkedIn is the authority on current employer), adds
the profile URL, adds new connections, and stamps `import_date`:

    mc -i Connections.csv --iformat linkedin -o data/contacts.json

Imports always fold into the existing `-o` DB (they never wipe it); manual edits
are preserved unless you pass `--force` to overwrite overlapping fields:

    mc -i archive.pst -o data/contacts.json           # adds to the existing DB
    mc -i archive.pst -o data/contacts.json --force   # let the import win on conflicts

Exclude whole domains while importing:

    mc -i archive.pst -o data/contacts.json --blacklist blacklist.txt

Dump a per-email JSONL corpus for an LLM (no contacts DB):

    mc -i mailbox.mbox -o emails.jsonl --llm

Export filtered contacts to a vCard:

    mc -i data/contacts.json --type customer,investor -o leads.vcf

Export filtered contacts to CSV:

    mc -i data/contacts.json --company Intel,AMD --min-emails 5 -o intel_amd.csv

Export only contacts at target-company domains (`--whitelist`), or drop
unwanted domains (`--blacklist`); both read one domain per line and ignore
`#` comments and blank lines, and match subdomains too. Each flag takes a
**list of files** (unioned), so you can keep categories in separate files:

    mc -i data/contacts.json --whitelist companies.txt -o targets.xlsx
    mc -i data/contacts.json --whitelist semiconductor.txt defense.txt equipment.txt -o targets.xlsx
    mc -i data/contacts.json --blacklist spam_domains.txt competitors.txt -o cleaned.json

Export in Outlook's column layout, as CSV or XLSX (`--oformat outlook`):

    mc -i data/contacts.json -o outlook.csv --oformat outlook
    mc -i data/contacts.json -o outlook.xlsx --oformat outlook

Deduplicate contacts sharing a first+last name, rewriting in place:

    mc -i data/contacts.json -o data/contacts.json --dedup

Merge one database into another (folding `extra.json` into `data/contacts.json`):

    mc -i extra.json -o data/contacts.json

## MC Help

```
usage: mc [-h] -i INPUT -o OUTPUT
          [--iformat {json,csv,xlsx,outlook,vcard,linkedin,mbox,pst,jsonl}]
          [--oformat {json,csv,xlsx,outlook,vcard,linkedin,mbox,pst,jsonl}] [--dedup]
          [--force] [--llm] [--max-body BYTES] [--include-cc] [--whitelist PATH [PATH ...]]
          [--blacklist PATH [PATH ...]] [--type TYPE] [--company COMPANY]
          [--first-name FIRST_NAME] [--last-name LAST_NAME]
          [--email-domain EMAIL_DOMAIN] [--min-emails MIN_EMAILS]
          [--max-emails MAX_EMAILS] [--min-sent MIN_SENT] [--max-sent MAX_SENT]
          [--min-received MIN_RECEIVED] [--max-received MAX_RECEIVED]
          [--last-after YYYY-MM-DD] [--last-before YYYY-MM-DD]
          [--first-after YYYY-MM-DD] [--first-before YYYY-MM-DD]

options:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        input path: a mailbox (.mbox/.pst), a vCard (.vcf/.vcd), an
                        Outlook CSV (--iformat outlook), or a contacts .json
  -o OUTPUT, --output OUTPUT
                        output path: a .json contacts DB, a .csv/.vcf export, or a .jsonl
                        corpus (with --llm)
  --iformat {json,csv,xlsx,outlook,vcard,linkedin,mbox,pst,jsonl}
                        force the input format instead of inferring it from the extension;
                        'outlook' reads an Outlook/Google CSV; 'linkedin' reads a LinkedIn
                        Connections export
  --oformat {json,csv,xlsx,outlook,vcard,linkedin,mbox,pst,jsonl}
                        force the output format instead of inferring it from the
                        extension; 'outlook' writes Outlook's CSV layout
  --dedup               merge contacts sharing a first+last name (json -> json)
  --force               when an imported record overlaps an existing one, overwrite the
                        existing text fields (company, title, name, ...) with the incoming
                        values; by default existing (hand-edited) values are kept
  --llm                 dump a per-email JSONL corpus (subject/from/to/date/body) from an
                        mbox/PST instead of building the DB
  --max-body BYTES      --llm: cap each message body to BYTES (default 262144; 0 =
                        unlimited) so attachment blobs do not dominate
  --include-cc          when importing a mailbox, also harvest the other To/Cc
                        recipients of mail you received (people on a thread with you,
                        not just the sender), counted as num_cc
  --whitelist PATH [PATH ...]
                        keep only contacts whose email domain matches an entry in these
                        files (one domain per line; '#' comments and blank lines ignored;
                        subdomains match too). Accepts multiple files (unioned) and may
                        be repeated.
  --blacklist PATH [PATH ...]
                        drop contacts/addresses whose email domain matches an entry in
                        these files (one per line; '#' comments and blank lines ignored;
                        subdomains match too). Accepts multiple files (unioned) and may
                        be repeated.
  --type TYPE           match contact type against any of LIST
                        (customer/competitor/investor/reporter/partner/vendor/other)
  --company COMPANY     match company against any of LIST
  --first-name FIRST_NAME
                        match first name against any of LIST
  --last-name LAST_NAME
                        match last name against any of LIST
  --email-domain EMAIL_DOMAIN
                        match primary email domain against any of LIST
  --min-emails MIN_EMAILS
                        minimum num_emails
  --max-emails MAX_EMAILS
                        maximum num_emails
  --min-sent MIN_SENT   minimum num_sent
  --max-sent MAX_SENT   maximum num_sent
  --min-received MIN_RECEIVED
                        minimum num_received
  --max-received MAX_RECEIVED
                        maximum num_received
  --last-after YYYY-MM-DD
                        last_interaction on or after this date
  --last-before YYYY-MM-DD
                        last_interaction on or before this date
  --first-after YYYY-MM-DD
                        first_interaction on or after this date
  --first-before YYYY-MM-DD
                        first_interaction on or before this date

```

## Database Record

The database is a JSON array of contact records. Each record has the same 16
fields, in this order:

```json
{
  "type": "customer",
  "friend": "",
  "last_name": "Vale",
  "first_name": "Jordan",
  "title": "CTO",
  "company": "Globex",
  "phone": "+16502530000",
  "address": "10 Loop, Springfield CA",
  "primary_email": "jordan@globex.com",
  "emails": ["jordan@globex.com", "jordan.vale@globex.com"],
  "num_emails": 50,
  "num_sent": 30,
  "num_received": 20,
  "num_cc": 0,
  "first_interaction": "2023-01-01",
  "last_interaction": "2025-03-15",
  "source": "work.mbox | takeout.mbox",
  "linkedin": "https://www.linkedin.com/in/jordanvale",
  "import_date": "2026-06-06"
}
```

| Field | Type | Description |
| --- | --- | --- |
| `type` | string | Annotation you fill in: one of `customer`, `competitor`, `investor`, `reporter`, `partner`, `vendor`, `other`. Blank on import. |
| `friend` | string | Annotation flag (e.g. `Y`); set from a vCard `friend` category, otherwise blank. |
| `last_name` | string | Surname, derived from the display name. |
| `first_name` | string | Given name, derived from the display name. |
| `title` | string | Annotation you fill in (job title); set from a vCard `TITLE`, otherwise blank. |
| `company` | string | Derived from the email domain; blank for free providers (gmail/yahoo/outlook/...). |
| `phone` | string | Extracted from the email signature (or a vCard `TEL`), normalized to `+E.164`; blank if none found. |
| `address` | string | Annotation you fill in; set from a vCard `ADR`/`LABEL`, otherwise blank. |
| `primary_email` | string | The most-used address; the record's key. |
| `emails` | string[] | All known addresses for the person, primary first. |
| `num_emails` | integer | Total messages involving this contact (`num_sent` + `num_received` + `num_cc`). |
| `num_sent` | integer | Messages you sent to this contact. |
| `num_received` | integer | Messages received from this contact. |
| `num_cc` | integer | Messages where they were a To/Cc co-recipient on a thread with you (0 unless imported with `--include-cc`). |
| `first_interaction` | string\|null | Earliest interaction date (`YYYY-MM-DD`), or `null` if unknown. |
| `last_interaction` | string\|null | Latest interaction date (`YYYY-MM-DD`), or `null` if unknown. |
| `source` | string | Origin file(s) the record came from, joined by ` \| `. |
| `linkedin` | string | LinkedIn profile URL; set by a LinkedIn import (`--iformat linkedin`), otherwise blank. |
| `import_date` | string | Date (`YYYY-MM-DD`) of the most recent non-database import (mbox/PST/vCard/Outlook/LinkedIn) that touched this record; blank for purely database-derived rows. |

The four annotation columns (`type`, `friend`, `title`, `address`) are left
blank on a mailbox import for you to fill in by hand; they are preserved across
re-imports and merges (see [Merge vs overwrite](#merge-vs-overwrite)). The same
fields are the columns of the CSV export, and map to the corresponding vCard
properties on export.


## Building the contacts database

An mbox/PST/vCard/Outlook-CSV input is treated as an **import**, building
contacts into the output database. A Gmail Takeout `.mbox`, an Outlook `.pst`,
and a vCard `.vcf`/`.vcd` are recognized by the `-i` extension. From a mailbox,
contacts are the people you have corresponded with (sent to or heard from), with
automated/bulk senders, spam, and nameless entries filtered out, identities
merged by display name, and company derived from the email domain.

```bash
mc -i "/path/to/Takeout/Mail/All mail Including Spam and Trash.mbox" -o data/contacts.json
mc -i "/path/to/archive.pst" -o data/contacts.json    # Outlook PST
mc -i "/path/to/contacts.vcf" -o data/contacts.json   # vCard (e.g. a Gmail export)
```

`-i` and `-o` are required. The output format follows the `-o` extension:
`.json`, `.csv`, and `.xlsx` are the interchangeable native database formats
(same columns, lossless round-trip -- edit the DB in Excel and re-import it), and
`.vcf` writes a vCard. Excel support is `.xlsx` only (via openpyxl); the legacy
binary `.xls` is not supported. To read an **Outlook-format CSV/XLSX** (the
column layout Outlook and Google Contacts export) pass `--iformat outlook`, since
a bare `.csv`/`.xlsx` is read as the native layout:

    mc -i contacts.csv --iformat outlook -o data/contacts.json

The Outlook reader takes First/Last Name, Job Title, Company, the E-mail Address
columns, Business Phone, and the business address columns; `type`/`friend` and
email counts are left blank/0. See `mc -h` for all options.

Importing a **vCard** adds its contacts directly (no message filtering): it maps
N/FN, ORG, TITLE, TEL, ADR, every EMAIL, and `CATEGORIES` (a category matching a
legal `type` value sets the type; a `friend` category sets the friend flag). Like
any import it folds into the existing database, so you can combine a vCard export
with an mbox-built database.

### What gets imported

A row is created for each **person you have corresponded with** -- anyone you
sent mail to or who sent mail to you. Specifically, an address is imported only
if **all** of these hold:

- It is a recipient (`To`/`Cc`) of mail you sent, **or** the sender (`From`) of
  mail you received (either direction qualifies). For PST, "sent" mail is the
  **Sent Items** folder. With `--include-cc`, the *other* `To`/`Cc` recipients of
  mail you received also qualify (people on a thread with you, counted as
  `num_cc`) -- broader reach, but noisier (large CC lists, mailing lists).
- The message is **not** spam: the Gmail `Spam` label, or for PST the **Junk
  Email** folder, is skipped.
- It is **not one of your own addresses** (auto-detected from the mbox
  `Delivered-To` header and the `From` of sent mail).
- It is **not an automated/bulk sender** -- e.g. `no-reply@`, `mailer-daemon`,
  `postmaster`, notifications, newsletters, marketing/unsubscribe addresses,
  `+`-tagged addresses (such as GitHub `reply+...`), or a bulk email-service /
  mailing-list domain (Mailchimp, SendGrid, Marketo, Beehiiv, GitHub, ...).
- Its domain is **not in `--blacklist`** (see below).
- The resulting contact has **both a first and last name** (single-name or
  org-style entries are dropped).

Then, across all imported addresses:

- Multiple addresses for the **same person** (matching display name) are merged
  into one row; the most-used address becomes the primary email.
- `phone` is pulled from the contact's **email signature** in mail they sent you
  (the signature region only -- bottom of the message / labeled lines). The
  most-frequently-seen number is kept, validated and normalized to `+E.164` via
  [phonenumbers](https://github.com/daviddrysdale/python-phonenumbers); numbers
  written without a country code are assumed US.
- `company` is derived from the email domain (blank for free providers like
  gmail/yahoo/outlook), and each row records sent/received counts, the first and
  last interaction dates, and the `source` filename (the `.mbox`/`.pst` it came
  from).

Pass `--blacklist PATH` to exclude whole domains from the contacts. The file
lists one domain per line (`#` comments and blank lines ignored); entries may be
written as `example.com` or `@example.com`, and subdomains are matched too:

```text
# blacklist.txt
recruiting-spam.com
@newsletters.example.org
```

`--whitelist PATH` is the inverse, used on **export**: it keeps only contacts
whose email domain (the `primary_email` or any address in `emails`) matches an
entry in the file, dropping everyone else. It uses the same file format as
`--blacklist` (one domain per line, `#` comments and blank lines ignored,
subdomains matched), so a categorized list with `# section` headers works as-is:

```text
# companies.txt
# -- semiconductor --
intel.com
nvidia.com
# -- agencies --
darpa.mil
```

`--whitelist` and `--blacklist` can be combined and apply to any export
(`-o .csv/.xlsx/.vcf/.json`); whitelist keeps matches, blacklist then removes
any that should still be dropped.

### Merging (the default) and `--force`

An import (and a DB&nbsp;->&nbsp;`.json` write) **always folds into the existing
output DB -- it never wipes it.** A missing output file is created fresh; an
existing one is read, merged into, and written back. There is no separate
"overwrite the whole file" mode: to start over, delete the file (or point `-o` at
a new path).

For an existing contact, the counts (`num_emails`, `num_sent`, `num_received`,
`num_cc`) are overwritten with the latest import, the email list is unioned, the
interaction date range widens, and `import_date` updates. Hand-edited text fields
(`type`/`friend`/`title`/`address`, plus name/company/phone) are **preserved** by
default -- the import only fills a blank. Pass **`--force`** to let the incoming
non-empty values **overwrite** those fields instead. Contacts present only in the
old file are always kept.

```bash
mc -i archive.pst -o data/contacts.json            # fold in; keep manual edits
mc -i archive.pst -o data/contacts.json --force    # let the import win on conflicts
mc -i extra.json  -o data/contacts.json            # fold one DB into another
```

This lets you re-run as a mailbox grows, or accumulate multiple sources, without
losing manual annotations. Imported rows include blank columns for you to fill in
by hand: `type` (one of `customer`, `competitor`, `investor`, `reporter`,
`partner`, `vendor`, `other`), `friend`, `title`, and `address`.

(Records are matched by email, or by LinkedIn profile URL when there is no email,
so email-less LinkedIn contacts survive a merge.)

### Importing from LinkedIn

A LinkedIn **Connections** export (`Settings -> Data privacy -> Get a copy of
your data -> Connections`) is the authority on a contact's *current* employer and
title. Import it with `--iformat linkedin` (the `.csv` extension alone is
ambiguous, so the format is explicit, like `outlook`); it folds into your existing
DB:

```bash
mc -i Connections.csv --iformat linkedin -o data/contacts.json
```

How it differs from a normal merge:

- **Matching:** by profile **URL**, then **email**, then normalized **first+last
  name** (LinkedIn omits most emails, so names do most of the work). A name that
  matches more than one existing contact is **skipped** (reported), not guessed.
- **Authority:** on a match, `company` and `title` are **overwritten** from
  LinkedIn (a normal merge would preserve them). The profile URL is stored in
  `linkedin`.
- **New connections are added** as contacts (most have no email -- they are
  identified by their LinkedIn URL). A connection with neither an email nor a URL
  is skipped (nothing to key it by).
- **`import_date`** is set to the date you run `mc` (use it later to reason about
  how fresh a contact's company is). It is stamped on every non-database import
  (mbox, PST, vCard, Outlook, LinkedIn), and left blank for database-only rows.

Re-running the same export is idempotent (URL/email matches refresh in place
rather than duplicating).

### Dump for an LLM

`--llm` skips the contacts database and instead writes a per-email **JSONL**
corpus (one JSON object per line) for feeding to an LLM. Each record is
`{subject, from, to, date, body}` with the full body, HTML stripped to text.
Every message is included except obvious `no-reply` senders, and it works on both
mbox and PST.

```bash
mc -i mailbox.mbox -o emails.jsonl --llm
mc -i archive.pst  -o emails.jsonl --llm    # works on PST too
```

The JSONL is streamed as messages are read, so it scales to very large mailboxes
without holding everything in memory.

## Exporting contacts

Giving a **JSON database** as `-i` with a `.csv`/`.vcf` output runs an
**export**: it selects a subset of contacts by per-column criteria and writes
the **whole record** for each match. The output format follows the `-o`
extension:

- **`.csv`** -- all database columns. Pass `--oformat outlook` to instead write
  Outlook's CSV column layout (`First Name`, `E-mail Address`, `Business Phone`,
  ...) that Outlook and Google Contacts import directly.
- **`.vcf`** -- a Gmail-compatible **vCard 3.0** file (importable into Google
  Contacts and Outlook), CRLF-delimited and line-folded to 75 octets.

Text filters take comma-separated lists (case-insensitive, match any); numeric
and date filters are inclusive ranges; all filters combine with AND. Note that
`--company` matches the derived company *name* (e.g. `Globex`), while
`--email-domain` matches the address domain (e.g. `globex.com`).

For a long list of domains, use `--whitelist FILE...` (keep only matches) and/or
`--blacklist FILE...` (drop matches) instead of a comma-separated `--email-domain`.
Unlike `--email-domain`, these read files (`#` comments and blank lines
ignored), match **any** of a contact's addresses (`primary_email` plus
`emails`), and match **subdomains** too (`intel.com` also catches
`fab.intel.com`). Each flag accepts **multiple files** (their domains are
unioned, and the flag may also be repeated), so you can keep each category in
its own file. They combine with each other and with all the column filters.

```bash
# Customers and investors (the `type` column you filled in) -> vCard:
mc -i data/contacts.json --type customer,investor -o leads.vcf

# Everyone at Intel or AMD with at least 5 emails, active since 2024 -> CSV:
mc -i data/contacts.json \
  --company Intel,AMD --min-emails 5 --last-after 2024-01-01 -o intel_amd.csv

# All intel.com contacts since 2025 -> vCard:
mc -i data/contacts.json \
  --email-domain intel.com --last-after 2025-01-01 -o intel.vcf

# Only contacts at target-company domains, split across category files -> XLSX:
mc -i data/contacts.json \
  --whitelist semiconductor.txt defense.txt equipment.txt -o targets.xlsx
```

The vCard maps name/emails (primary marked `PREF`), `company`->ORG,
`title`->TITLE, `phone`->TEL, `address`->ADR/LABEL, and `type`/`friend`->
CATEGORIES, plus a NOTE with the email counts and last-contact date.

`--type` accepts only the legal values (`customer`, `competitor`, `investor`,
`reporter`, `partner`, `vendor`, `other`). See `mc -h` for the full set of
filters (`--type`, `--first-name`, `--last-name`, `--email-domain`,
`--whitelist`, `--blacklist`, `--min/max-emails`, `--min/max-sent`,
`--min/max-received`, `--first-after/before`, `--last-after/before`).

## Deduplicating contacts

A JSON-to-JSON run with `--dedup` merges rows that share the same first **and**
last name (case-insensitive), which can accumulate from multiple imports or
manual edits.

```bash
mc -i data/contacts.json -o data/contacts.json --dedup   # in place
mc -i data/contacts.json -o deduped.json --dedup          # or to a new file
```

When duplicates are merged:

- counts (`num_emails`/`num_sent`/`num_received`/`num_cc`) are **summed**, `emails`
  and `source` are **unioned**, and the interaction date range **widens**;
- conflicting annotation fields (`type`, `friend`, `title`, `company`, `phone`,
  `address`) are **kept all** (distinct values joined with ` | `), so no manual
  edit is lost;
- the `primary_email` and name casing come from the highest-volume duplicate.

Rows missing a first or last name are left untouched. `-o` may equal `-i` to
rewrite in place. Dedup is a JSON-to-JSON operation; to deduplicate before an
export, dedup to `.json` first and then export. Because matching is by name
only, two different people with the same name will be merged -- the joined
`company` and multiple `emails` make such cases easy to spot for manual review.
