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

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## MC Program

The MailCompiler command line utility is called 'mc'

```
mc -h                                        print out command line help
mc -i MBOX|PST|VCF|CSV -o DB.json [...]      import contacts into a JSON DB
mc -i DB.json -o OUT.{csv,vcf} [filters]     export matching records
mc -i DB.json -o OUT.json --dedup            merge same-name contacts
```

## Examples

Build the JSON contacts database from a Gmail Takeout mbox:

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
  "first_interaction": "2023-01-01",
  "last_interaction": "2025-03-15",
  "source": "work.mbox | takeout.mbox"
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
| `num_emails` | integer | Total messages exchanged (`num_sent` + `num_received`). |
| `num_sent` | integer | Messages you sent to this contact. |
| `num_received` | integer | Messages received from this contact. |
| `first_interaction` | string\|null | Earliest interaction date (`YYYY-MM-DD`), or `null` if unknown. |
| `last_interaction` | string\|null | Latest interaction date (`YYYY-MM-DD`), or `null` if unknown. |
| `source` | string | Origin file(s) the record came from, joined by ` \| `. |

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
