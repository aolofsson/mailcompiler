![MailCompiler](docs/postverket.svg)

MailCompiler is a utility for reclaiming control over emails and contacts via
automatic imports from .mbox and .pst exports into a human readableJ JSON
database.

## Motivation
- You need to own your our data!!
- Scraping your own inbox to put together outreach lists is a waste of life
- CRMs are awful
- Email client import/output is barely functional
- Email client search functions are completely broken

## Key features

- Import from Gmail Takeout `.mbox` and Outlook `.pst`(handles 20 GB+)
- Automatic extraction of email conversations into contacts
- Automatic extraction of phone numbers from email signatures
- Import/export support for VCARD (3.0) and CSV (Outlook) contact lists
- Streaming .mbox into LLM frindly JSONL corpus
- Incremental non-destructive merging of multiple JSON databases
- Deduplication of records
- Black list email list support
- Automatic filtering of bot farm email addresses
- Record filtering on export

## Install (development)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

`mc` has one subcommand per operation, all with a uniform `-i/--input`:

```
mc import  -i MBOX|PST|VCF|CSV -o OUT [...]   build/merge the contacts DB
mc export  -i CONTACTS -o OUT.{csv,vcf}       export matching records
mc dedup   -i CONTACTS -o OUT                 merge same-name contacts
```

Import reads a Gmail Takeout `.mbox`, an Outlook `.pst`, a vCard `.vcf`/`.vcd`,
or (with `--outlook`) an Outlook CSV -- chosen by file extension. The database is
JSON (the native format); pass `-o something.csv` for a lossless CSV copy.
`mc import --llm` instead writes a per-email JSONL corpus for LLMs.

Without installing, run it as a module: `python -m mailcompiler.mailcompiler <command> ...`.

## Building the contacts database

`mc import` reads a Gmail Takeout `.mbox`, an Outlook `.pst`, or a vCard
`.vcf`/`.vcd` (chosen by the `-i` file extension) into `contacts.json`. From a
mailbox, contacts are the people you have corresponded with (sent to or heard
from), with automated/bulk senders, spam, and nameless entries filtered out,
identities merged by display name, and company derived from the email domain.

```bash
mc import -i "/path/to/Takeout/Mail/All mail Including Spam and Trash.mbox" -o data
mc import -i "/path/to/archive.pst" -o data       # Outlook PST
mc import -i "/path/to/contacts.vcf" -o data      # vCard (e.g. a Gmail export)
```

Importing a **vCard** adds its contacts directly (no message filtering): it maps
N/FN, ORG, TITLE, TEL, ADR, every EMAIL, and `CATEGORIES` (a category matching a
legal `type` value sets the type; a `friend` category sets the friend flag).
Email counts default to 0 and a contact with no email address is skipped. This
merges into the database like any other import, so you can fold a vCard export
into an mbox-built database.

Add `--outlook` to import an **Outlook-format CSV** (the column layout Outlook
and Google Contacts export): `mc import --outlook -i contacts.csv -o data`. It
reads First/Last Name, Job Title, Company, the E-mail Address columns, Business
Phone, and the business address columns; `type`/`friend` and email counts are
left blank/0.

`-i` and `-o` are required. The output format follows the `-o` extension:
`.json` is the native database, `.csv` exports CSV; a directory (as above)
writes `contacts.json` inside it. See `mc import -h` for all options.

PST reading uses [pypff](https://github.com/libyal/libpff) (`libpff-python`),
installed automatically with the package.

### What gets imported

A row is created for each **person you have corresponded with** -- anyone you
sent mail to or who sent mail to you. Specifically, an address is imported only
if **all** of these hold:

- It is a recipient (`To`/`Cc`) of mail you sent, **or** the sender (`From`) of
  mail you received (either direction qualifies). For PST, "sent" mail is the
  **Sent Items** folder.
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

### Merge vs overwrite

If the output file already exists, the importer **merges** into it (in whichever
format the output path uses). For an existing contact, the counts (`num_emails`,
`num_sent`, `num_received`) are overwritten with the latest import, the email
list is unioned, and the interaction date range widens; hand-edited fields
(the blank annotation columns `type`/`friend`/`title`/`address`, plus
name/company) are preserved. Contacts present only in the old file are left
untouched. This lets you re-run as a mailbox grows, or accumulate multiple
mailboxes, without losing manual annotations.

Imported rows include blank columns for you to fill in by hand: `type` (one of
`customer`, `competitor`, `investor`, `reporter`, `partner`, `vendor`, `other`),
`friend`, `title`, and `address`.

Use `-f` / `--force` to ignore the existing file and write a fresh one instead
(this discards any manual edits such as `type`).

### Dump for an LLM

`mc import --llm` skips the contacts database and instead writes a per-email
**JSONL** corpus (one JSON object per line) for feeding to an LLM. Each record is
`{subject, from, to, date, body}` with the full body, HTML stripped to text.
Every message is included except obvious `no-reply` senders, and it works on both
mbox and PST.

```bash
mc import --llm -i mailbox.mbox -o emails.jsonl
mc import --llm -i archive.pst  -o data            # writes data/emails.jsonl
```

The JSONL is streamed as messages are read, so it scales to very large mailboxes
without holding everything in memory.

## Exporting contacts

`mc export` selects a subset of contacts by per-column criteria and writes the
**whole record** for each match. The output format follows the `-o` extension
(required):

- **`.csv`** -- all database columns. Add `--outlook` to instead write Outlook's
  CSV column layout (`First Name`, `E-mail Address`, `Business Phone`, ...) that
  Outlook and Google Contacts import directly.
- **`.vcf`** -- a Gmail-compatible **vCard 3.0** file (importable into Google
  Contacts and Outlook), CRLF-delimited and line-folded to 75 octets.

Text filters take comma-separated lists (case-insensitive, match any); numeric
and date filters are inclusive ranges; all filters combine with AND. Note that
`--company` matches the derived company *name* (e.g. `Globex`), while
`--email-domain` matches the address domain (e.g. `globex.com`).

```bash
# Customers and investors (the `type` column you filled in) -> vCard:
mc export -i data/contacts.json --type customer,investor -o leads.vcf

# Everyone at Intel or AMD with at least 5 emails, active since 2024 -> CSV:
mc export -i data/contacts.json \
  --company Intel,AMD --min-emails 5 --last-after 2024-01-01 -o intel_amd.csv

# All intel.com contacts since 2025 -> vCard:
mc export -i data/contacts.json \
  --email-domain intel.com --last-after 2025-01-01 -o intel.vcf
```

The vCard maps name/emails (primary marked `PREF`), `company`->ORG,
`title`->TITLE, `phone`->TEL, `address`->ADR/LABEL, and `type`/`friend`->
CATEGORIES, plus a NOTE with the email counts and last-contact date.

`--type` accepts only the legal values (`customer`, `competitor`, `investor`,
`reporter`, `partner`, `vendor`, `other`). See `mc export -h` for the full set of
filters (`--type`, `--first-name`, `--last-name`, `--min/max-emails`,
`--min/max-sent`, `--min/max-received`, `--first-after/before`,
`--last-after/before`).

## Deduplicating contacts

`mc dedup` merges rows that share the same first **and** last name
(case-insensitive), which can accumulate from multiple imports or manual edits.

```bash
mc dedup -i data/contacts.json -o data/contacts.json   # in place
mc dedup -i data/contacts.json -o deduped.csv           # or to a new file
```

When duplicates are merged:

- counts (`num_emails`/`num_sent`/`num_received`) are **summed**, `emails` and
  `source` are **unioned**, and the interaction date range **widens**;
- conflicting annotation fields (`type`, `friend`, `title`, `company`, `phone`,
  `address`) are **kept all** (distinct values joined with ` | `), so no manual
  edit is lost;
- the `primary_email` and name casing come from the highest-volume duplicate.

Rows missing a first or last name are left untouched. `-o` may equal `-i` to
rewrite in place, and the output format follows the `-o` extension (JSON/CSV).
Because matching is by name only, two different people with the same name will
be merged -- the joined `company` and multiple `emails` make such cases easy to
spot for manual review.
