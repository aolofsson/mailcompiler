![MailCompiler](docs/postverket.svg)

MailCompiler is a simple Python program for: importing data from an email archives (.mbox and .pst) into into a human readable local JSON file/databsae and for converting between different contact file formats (VCARD, CSV).

## Motivation
- You need to own your data!
- Manually scraping inboxes to put together outreach lists is a waste of life
- CRMs are over engineered for startups
- Email client import/output is barely functional
- Email client search functions are completely broken

## Key features
- Clean human readable JSON contact database
- Fast import from Gmail Takeout `.mbox` and Outlook `.pst` (handles 20 GB+)
- Import/export support for VCARD (3.0) and CSV (Outlook) contact lists
- Streaming .mbox into LLM friendly JSON corpus
- Automatic extraction of email conversations into contacts
- Automatic extraction of phone numbers from email signatures
- Incremental non-destructive merging of multiple JSON databases
- Deduplication of records
- Black list email list support
- Automatic filtering of bot farm email addresses
- Record filtering on export

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## MC Email Utility

The MailCompiler command line utility is called 'mc'

`mc` has no subcommands: the operation is inferred from the `-i`/`-o` file
formats (taken from the extension, or forced with `--iformat`/`--oformat`). A
mailbox, vCard, or Outlook CSV input **imports** into a contacts database; a
JSON database input **exports** (to `.csv`/`.vcf`) or, with `--dedup`,
**deduplicates** (to `.json`). JSON is the native database; CSV and vCard are
interchange formats.

```
mc -i MBOX|PST|VCF|CSV -o DB.json [...]      import contacts into a JSON DB
mc -i DB.json -o OUT.{csv,vcf} [filters]     export matching records
mc -i DB.json -o OUT.json --dedup            merge same-name contacts
```

## Examples

Build the contacts DB from a Gmail Takeout mbox:

    mc -i "All mail Including Spam and Trash.mbox" -o data/contacts.json

Import an Outlook PST instead:

    mc -i archive.pst -o data/contacts.json

Import a vCard export (Google Contacts / Gmail):

    mc -i contacts.vcf -o data/contacts.json

Import an Outlook / Google Contacts CSV export (`--iformat outlook`):

    mc -i contacts.csv --iformat outlook -o data/contacts.json

Merge a new import into an existing DB, preserving manual edits:

    mc -i archive.pst -o data/contacts.json --merge

Exclude whole domains while importing:

    mc -i archive.pst -o data/contacts.json --blacklist blacklist.txt

Dump a per-email JSONL corpus for an LLM (no contacts DB):

    mc -i mailbox.mbox -o emails.jsonl --llm

Export filtered contacts to a vCard:

    mc -i data/contacts.json --type customer,investor -o leads.vcf

Export filtered contacts to CSV:

    mc -i data/contacts.json --company Intel,AMD --min-emails 5 -o intel_amd.csv

Export in Outlook's CSV column layout (`--oformat outlook`):

    mc -i data/contacts.json -o outlook.csv --oformat outlook

Deduplicate contacts sharing a first+last name, rewriting in place:

    mc -i data/contacts.json -o data/contacts.json --dedup

Merge one database into another (folding `extra.json` into `data/contacts.json`):

    mc -i extra.json -o data/contacts.json --merge

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
`.json` is the native database, `.csv` writes the full-column CSV, and `.vcf`
writes a vCard. To import an **Outlook-format CSV** (the column layout Outlook
and Google Contacts export) pass `--iformat outlook`, since a bare `.csv` is
read as the native CSV layout:

    mc -i contacts.csv --iformat outlook -o data/contacts.json

The Outlook reader takes First/Last Name, Job Title, Company, the E-mail Address
columns, Business Phone, and the business address columns; `type`/`friend` and
email counts are left blank/0. See `mc -h` for all options.

Importing a **vCard** adds its contacts directly (no message filtering): it maps
N/FN, ORG, TITLE, TEL, ADR, every EMAIL, and `CATEGORIES` (a category matching a
legal `type` value sets the type; a `friend` category sets the friend flag). With
`--merge` this folds into an existing database like any other import, so you can
combine a vCard export with an mbox-built database.

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

By default an import **overwrites** the output file. Pass `--merge` to fold the
import into an existing database instead (in whichever format the output path
uses). For an existing contact, the counts (`num_emails`, `num_sent`,
`num_received`) are overwritten with the latest import, the email list is
unioned, and the interaction date range widens; hand-edited fields (the blank
annotation columns `type`/`friend`/`title`/`address`, plus name/company) are
preserved. Contacts present only in the old file are left untouched. This lets
you re-run as a mailbox grows, or accumulate multiple mailboxes, without losing
manual annotations:

```bash
mc -i archive.pst -o data/contacts.json --merge   # fold into the existing DB
```

Imported rows include blank columns for you to fill in by hand: `type` (one of
`customer`, `competitor`, `investor`, `reporter`, `partner`, `vendor`, `other`),
`friend`, `title`, and `address`. `--merge` also folds one JSON database into
another (`mc -i extra.json -o data/contacts.json --merge`).

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

```bash
# Customers and investors (the `type` column you filled in) -> vCard:
mc -i data/contacts.json --type customer,investor -o leads.vcf

# Everyone at Intel or AMD with at least 5 emails, active since 2024 -> CSV:
mc -i data/contacts.json \
  --company Intel,AMD --min-emails 5 --last-after 2024-01-01 -o intel_amd.csv

# All intel.com contacts since 2025 -> vCard:
mc -i data/contacts.json \
  --email-domain intel.com --last-after 2025-01-01 -o intel.vcf
```

The vCard maps name/emails (primary marked `PREF`), `company`->ORG,
`title`->TITLE, `phone`->TEL, `address`->ADR/LABEL, and `type`/`friend`->
CATEGORIES, plus a NOTE with the email counts and last-contact date.

`--type` accepts only the legal values (`customer`, `competitor`, `investor`,
`reporter`, `partner`, `vendor`, `other`). See `mc -h` for the full set of
filters (`--type`, `--first-name`, `--last-name`, `--min/max-emails`,
`--min/max-sent`, `--min/max-received`, `--first-after/before`,
`--last-after/before`).

## Deduplicating contacts

A JSON-to-JSON run with `--dedup` merges rows that share the same first **and**
last name (case-insensitive), which can accumulate from multiple imports or
manual edits.

```bash
mc -i data/contacts.json -o data/contacts.json --dedup   # in place
mc -i data/contacts.json -o deduped.json --dedup          # or to a new file
```

When duplicates are merged:

- counts (`num_emails`/`num_sent`/`num_received`) are **summed**, `emails` and
  `source` are **unioned**, and the interaction date range **widens**;
- conflicting annotation fields (`type`, `friend`, `title`, `company`, `phone`,
  `address`) are **kept all** (distinct values joined with ` | `), so no manual
  edit is lost;
- the `primary_email` and name casing come from the highest-volume duplicate.

Rows missing a first or last name are left untouched. `-o` may equal `-i` to
rewrite in place. Dedup is a JSON-to-JSON operation; to deduplicate before an
export, dedup to `.json` first and then export. Because matching is by name
only, two different people with the same name will be merged -- the joined
`company` and multiple `emails` make such cases easy to spot for manual review.
