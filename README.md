![MailCompiler](docs/postverket.svg)

**MailCompiler** (`mc`) is a single-command CLI that lets you mine your mailboxes and address books (Gmail Takeout, Outlook archives, vCard, LinkedIn) and consolidate all of the data into a single, unified, de-duplicated, human-readable JSON contact database that you control.

> Own your inbox. Own your network.

![MailCompiler Architecture](docs/mc_arch.png)

## Motivation
- Your contact network is valuable, don't lose touch!
- Don't let platforms control you, own your data!
- Email client search functions are basically useless, we need a LLM friendly DB.
- Manually scraping old email inboxes to recover contacts is a waste of life.

## Key features
- Clean, human-readable JSON contact database
- Fast import from Gmail Takeout `.mbox` and Outlook `.pst` (handles 20 GB+)
- Import/export support for vCard 3.0 and CSV (Outlook) contact lists
- Lossless import/export between xlsx/csv files and the JSON database
- Streaming `.mbox` into an LLM-friendly JSON corpus
- Automatic extraction of email conversations into contacts
- Optional, opt-in phone-number mining from email signatures (`--discover-phones`)
- Incremental, non-destructive merging into a JSON database
- Deduplication and reconciliation of records
- Automatic industry categorization from a built-in [yellowpages directory](mailcompiler/yellowpages.csv)
- Blacklist support for email domains
- Automatic filtering of bot-farm email addresses
- Record filtering on export

## TL;DR

`mc` has four commands -- `import`, `export`, `reconcile`, `scrape` -- over one
lossless JSON database, set once via `--db` or `$MC_DB`.

```bash
pip install mailcompiler                       # install the `mc` command
export MC_DB=~/contacts.json                   # where the database lives
mc import takeout.mbox                         # import mbox contacts
mc import connections.csv --format linkedin    # merge in LinkedIn connections
mc reconcile                                   # clean + merge duplicates
mc export profit.xlsx                          # export to Excel
```

## Examples

The database is `$MC_DB` (or `--db PATH`). `import` folds a source into it;
`export` writes a filtered view out of it; `reconcile` cleans it in place.

Build the database from a Gmail Takeout mbox (into `$MC_DB`):

    mc import "All mail Including Spam and Trash.mbox"

Import into an explicit database file instead of `$MC_DB`:

    mc import "All mail Including Spam and Trash.mbox" --db data/contacts.json

Import an Outlook PST:

    mc import archive.pst

Import a vCard export (Google Contacts / Gmail):

    mc import contacts.vcf

Import an Outlook / Google Contacts CSV export (`.csv` is ambiguous, so name it):

    mc import contacts.csv --format outlook

Enrich from a LinkedIn Connections export (overwrites company/title -- LinkedIn
is the authority on current employer -- adds the profile URL, adds new
connections, and stamps `import_date`):

    mc import Connections.csv --format linkedin

Imports always fold into the database (they never wipe it); manual edits are
preserved unless you pass `--force` to overwrite overlapping fields:

    mc import archive.pst                       # adds to the database
    mc import archive.pst --force               # let the import win on conflicts

Exclude whole domains while importing:

    mc import archive.pst --blacklist blacklist.txt

Export the database to an Excel spreadsheet:

    mc export contacts.xlsx

Export filtered contacts to a vCard:

    mc export leads.vcf --category "Semiconductor Devices,Defense"

Export filtered contacts to CSV:

    mc export intel_amd.csv --company Intel,AMD --min-emails 5

Export only contacts at target-company domains (`--whitelist`), or drop
unwanted domains (`--blacklist`); both read one domain per line and ignore
`#` comments and blank lines, and match subdomains too. Each flag takes a
**list of files** (unioned), so you can keep categories in separate files:

    mc export targets.xlsx --whitelist semiconductor.txt defense.txt
    mc export cleaned.json --blacklist spam_domains.txt competitors.txt

Export in Outlook's column layout, as CSV or XLSX (`--format outlook`):

    mc export outlook.csv  --format outlook
    mc export outlook.xlsx --format outlook

Clean and merge duplicate records, in place:

    mc reconcile

Merge one database into another (folding `extra.json` into the database):

    mc import extra.json --db data/contacts.json

Scrape a mailbox into a per-email JSONL corpus for an LLM (no database):

    mc scrape mailbox.mbox -o emails.jsonl

### Set the database once with `$MC_DB`

So you don't repeat `--db contacts.json` on every command, point `$MC_DB` at your
database; `import`/`export`/`reconcile` then use it automatically (`--db` still wins):

```bash
export MC_DB=~/contacts.json     # in your shell rc
mc import takeout.mbox           # folds into $MC_DB (created if absent)
mc import archive.pst            # fold another source in
mc reconcile                     # cleans $MC_DB in place
mc export leads.xlsx             # exports from $MC_DB
```


## MC Help

`mc` is split into subcommands; run `mc <command> -h` for any one.

```
$ mc -h
usage: mc [-h] {import,export,reconcile,scrape} ...

mailcompiler: aggregate, de-duplicate and query a contacts database. The
database is set by --db or $MC_DB.

positional arguments:
  {import,export,reconcile,scrape}
    import              fold a source into the database
    export              write a filtered view of the database to a file
    reconcile           clean + merge duplicates in the database, in place
    scrape              scrape a mailbox into a per-email JSONL corpus

options:
  -h, --help            show this help message and exit

$ mc import -h
usage: mc import [-h] [-v] [--db PATH]
                 [--format {mbox,pst,vcard,outlook,linkedin,json,csv,xlsx}]
                 [--force] [--no-cc] [--discover-phones] [--self-phone LIST]
                 [--blacklist PATH [PATH ...]]
                 source

positional arguments:
  source                file to import (mbox/PST/vCard/Outlook/LinkedIn CSV,
                        or a json/csv/xlsx DB)

options:
  -h, --help            show this help message and exit
  -v, --verbose         print every discard/action to stderr
  --db PATH             contacts database path (.json/.csv/.xlsx); defaults to
                        $MC_DB
  --format {mbox,pst,vcard,outlook,linkedin,json,csv,xlsx}
                        override format inference; use 'outlook' or 'linkedin'
                        for a .csv
  --force               overwrite existing text fields with incoming values
                        (default keeps hand-edited values)
  --no-cc               mailbox import: skip the other To/Cc recipients of
                        mail you received (less noise)
  --discover-phones     (OFF by default) mine phone numbers from the signature
                        region of received mail and credit the most-frequent
                        number to the sender. HEURISTIC AND ERROR-PRONE: a
                        signature quoted at the bottom of a reply thread is
                        easily misattributed, so ~15%+ of discovered numbers
                        are wrong (one number can spread across many
                        contacts). Structured imports (vCard/Outlook/LinkedIn)
                        keep their phone fields regardless. Pair with --self-
                        phone; reconcile drops numbers shared across many
                        contacts.
  --self-phone LIST     your own phone number(s) (comma-separated; also read
                        from $MC_SELF_PHONE) to never ingest as a contact's
                        number
  --blacklist PATH [PATH ...]
                        drop addresses whose domain matches an entry in these
                        files (subdomains match; repeatable)

$ mc export -h
usage: mc export [-h] [-v] [--db PATH]
                 [--format {json,csv,xlsx,outlook,vcard}]
                 [--category CATEGORY] [--company COMPANY]
                 [--first-name FIRST_NAME] [--last-name LAST_NAME]
                 [--email-domain EMAIL_DOMAIN] [--min-emails MIN_EMAILS]
                 [--max-emails MAX_EMAILS] [--min-sent MIN_SENT]
                 [--max-sent MAX_SENT] [--min-received MIN_RECEIVED]
                 [--max-received MAX_RECEIVED] [--min-ranking MIN_RANKING]
                 [--max-ranking MAX_RANKING] [--last-after YYYY-MM-DD]
                 [--last-before YYYY-MM-DD] [--first-after YYYY-MM-DD]
                 [--first-before YYYY-MM-DD] [--whitelist PATH [PATH ...]]
                 [--blacklist PATH [PATH ...]]
                 dest

positional arguments:
  dest                  output file (.csv/.xlsx/.vcf/.json; format from the
                        extension)

options:
  -h, --help            show this help message and exit
  -v, --verbose         print every discard/action to stderr
  --db PATH             contacts database path (.json/.csv/.xlsx); defaults to
                        $MC_DB
  --format {json,csv,xlsx,outlook,vcard}
                        override output format; 'outlook' writes Outlook's
                        column layout
  --category CATEGORY   match contact category against any of LIST
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
  --min-ranking MIN_RANKING
                        minimum ranking (0-100)
  --max-ranking MAX_RANKING
                        maximum ranking (0-100)
  --last-after YYYY-MM-DD
                        last_interaction on or after this date
  --last-before YYYY-MM-DD
                        last_interaction on or before this date
  --first-after YYYY-MM-DD
                        first_interaction on or after this date
  --first-before YYYY-MM-DD
                        first_interaction on or before this date
  --whitelist PATH [PATH ...]
                        keep only contacts whose email domain matches an entry
                        in these files (subdomains match; repeatable)
  --blacklist PATH [PATH ...]
                        drop contacts whose email domain matches an entry in
                        these files (subdomains match; repeatable)

$ mc reconcile -h
usage: mc reconcile [-h] [-v] [--db PATH] [--self-phone LIST]

options:
  -h, --help         show this help message and exit
  -v, --verbose      print every discard/action to stderr
  --db PATH          contacts database path (.json/.csv/.xlsx); defaults to
                     $MC_DB
  --self-phone LIST  your own phone number(s) (comma-separated; also read from
                     $MC_SELF_PHONE) to never ingest as a contact's number

$ mc scrape -h
usage: mc scrape [-h] [-v] -o PATH [--db PATH] [--format {mbox,pst}]
                 [--max-body BYTES]
                 source

positional arguments:
  source                mbox/PST file

options:
  -h, --help            show this help message and exit
  -v, --verbose         print every discard/action to stderr
  -o PATH, --output PATH
                        output .jsonl path
  --db PATH             contacts database (.json/.csv/.xlsx; default $MC_DB)
                        used to set each record's contact_id; optional
  --format {mbox,pst}   override format inference
  --max-body BYTES      cap each message body to BYTES (default 262144; 0 =
                        unlimited)
```


## Database Record

The database is a JSON array of contact records. Each record has the same 24
fields, in this order:

```json
{
  "last_name": "Vale",
  "first_name": "Jordan",
  "friend": true,
  "title": "CTO",
  "company": "Globex",
  "category": "Semiconductor Devices",
  "primary_phone": "+16502530000",
  "phone_numbers": ["+16502530000", "+14155550000"],
  "address": "10 Loop, Springfield CA",
  "birthday": "1985-04-12",
  "primary_email": "jordan@globex.com",
  "emails": ["jordan@globex.com", "jordan.vale@globex.com"],
  "num_emails": 50,
  "num_sent": 30,
  "num_received": 20,
  "first_interaction": "2023-01-01",
  "last_interaction": "2025-03-15",
  "source": "work.mbox | takeout.mbox",
  "linkedin": "https://www.linkedin.com/in/jordanvale",
  "github": "https://github.com/jvale",
  "ranking": 75,
  "notes": "Met at the 2025 conference.",
  "import_date": "2026-06-06",
  "id": "11111111-1111-1111-1111-111111111111"
}
```

| Field | Type | Description |
| --- | --- | --- |
| `last_name` | string | Surname, derived from the display name. |
| `first_name` | string | Given name, derived from the display name. |
| `friend` | boolean | Annotation flag; set from a vCard `friend` category. On import the values `Y`/`1`/`true`/`yes` (case-insensitive) become `true`, everything else `false`. |
| `title` | string | Annotation you fill in (job title); set from a vCard `TITLE`, otherwise blank. |
| `company` | string | Derived from the email domain; blank for free providers (gmail/yahoo/outlook/...). |
| `category` | string | Industry segment (e.g. `Semiconductor Devices`, `Defense`, `Venture Capital`, `Academic`), set automatically from the bundled yellowpages directory by email domain during `--reconcile`. Blank on import and for unlisted domains. |
| `primary_phone` | string | The preferred phone number, normalized to `+E.164`; blank if none found. |
| `phone_numbers` | string[] | All known numbers, primary first. From vCard `TEL` / Outlook columns, or (only with `--discover-phones`) mined from mail signatures. |
| `address` | string | Annotation you fill in; set from a vCard `ADR`/`LABEL`, otherwise blank. |
| `birthday` | string | `YYYY-MM-DD`; annotation you fill in, or from a vCard `BDAY` / Outlook Birthday. Blank if unknown. |
| `primary_email` | string | The most-used address; the merge key. |
| `emails` | string[] | All known addresses for the person, primary first. |
| `num_emails` | integer | Total direct messages exchanged (`num_sent` + `num_received`). 0 for a contact you only shared a thread with. |
| `num_sent` | integer | Messages you sent to this contact. |
| `num_received` | integer | Messages received from this contact. |
| `first_interaction` | string\|null | Earliest interaction date (`YYYY-MM-DD`), or `null` if unknown. |
| `last_interaction` | string\|null | Latest interaction date (`YYYY-MM-DD`), or `null` if unknown. |
| `source` | string | Origin file(s) the record came from, joined by ` \| `. |
| `linkedin` | string | LinkedIn profile URL; set by a LinkedIn import (`--format linkedin`), otherwise blank. |
| `github` | string | GitHub profile URL/handle; annotation you fill in, or from a vCard `URL`. Blank otherwise. |
| `ranking` | integer | User defined contact ranking `0`-`100` (default `0`); filterable with `--min-ranking`/`--max-ranking`. Merges keep the higher value. |
| `notes` | string | Free-text annotation you fill in. Blank otherwise. |
| `import_date` | string | Date (`YYYY-MM-DD`) of the most recent non-database import (mbox/PST/vCard/Outlook/LinkedIn) that touched this record; blank for purely database-derived rows. |
| `id` | string | Stable per-record UUID, minted when the record is first created and preserved across merges. |

The annotation columns (`friend`, `title`, `address`, `birthday`, `github`,
`ranking`, `notes`) are left blank on a mailbox import for you to fill in by
hand; they are preserved across re-imports and merges (see
[Merging and `--force`](#merging-the-default-and---force)).
(`category` is set automatically by `--reconcile`, not hand-filled.) The same
fields are the columns of the CSV export, and map to the corresponding vCard
properties on export.


## Building the contacts database

`mc import SOURCE` folds a mailbox/vCard/Outlook-CSV/LinkedIn source into the
database. A Gmail Takeout `.mbox`, an Outlook `.pst`, and a vCard `.vcf`/`.vcd`
are recognized by the `SOURCE` extension. From a mailbox, contacts are the people
you have corresponded with (sent to or heard from), with automated/bulk senders,
spam, and nameless entries filtered out, identities merged by display name, and
company derived from the email domain.

```bash
mc import "/path/to/Takeout/Mail/All mail Including Spam and Trash.mbox" --db data/contacts.json
mc import "/path/to/archive.pst"  --db data/contacts.json    # Outlook PST
mc import "/path/to/contacts.vcf" --db data/contacts.json    # vCard (e.g. a Gmail export)
```

The database is `--db PATH`, or `$MC_DB` when `--db` is omitted (see
[Set the database once with `$MC_DB`](#set-the-database-once-with-mc_db)). The
database format follows the `--db` extension:
`.json`, `.csv`, and `.xlsx` are the interchangeable native database formats
(same columns, lossless round-trip -- edit the DB in Excel and re-import it).
Excel support is `.xlsx` only (via openpyxl); the legacy binary `.xls` is not
supported. To read an **Outlook-format CSV/XLSX** (the column layout Outlook and
Google Contacts export) pass `--format outlook`, since a bare `.csv`/`.xlsx` is
read as the native layout:

    mc import contacts.csv --format outlook --db data/contacts.json

The Outlook reader takes First/Last Name, Job Title, Company, the E-mail Address
columns, Business Phone, and the business address columns; `category`/`friend` and
email counts are left blank/0. See `mc -h` for all options.

Importing a **vCard** adds its contacts directly (no message filtering): it maps
N/FN, ORG, TITLE, TEL, ADR, every EMAIL, and `CATEGORIES` (the first non-`friend`
category becomes the contact's `category`; a `friend` category sets the friend
flag). Like any import it folds into the existing database, so you can combine a
vCard export with an mbox-built database.

### What gets imported

A row is created for each **person you have corresponded with** -- anyone you
sent mail to or who sent mail to you. Specifically, an address is imported only
if **all** of these hold:

- It appeared on a message with you: a recipient (`To`/`Cc`) of mail you sent
  (counts as `num_sent`), the sender (`From`) of mail you received (`num_received`),
  **or** one of the *other* `To`/`Cc` recipients of mail you received -- people on
  a thread with you, even if you never corresponded directly (these count 0
  sent/received). For PST, "sent" mail is the **Sent Items** folder. Including
  thread co-recipients gives broad reach but is noisier -- large CC lists,
  mailing lists -- which the bot/no-reply/blacklist filters below help trim;
  pass **`--no-cc`** to skip them entirely and keep only direct correspondents.
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
- `phone_numbers` are **not** mined from mail by default. With **`--discover-phones`**
  (opt-in) they are pulled from the contact's **email signature** in mail they
  sent you (the signature region only -- bottom of the message / labeled lines),
  validated and normalized to `+E.164` via
  [phonenumbers](https://github.com/daviddrysdale/python-phonenumbers) (numbers
  written without a country code are assumed US); the most-frequently-seen
  number becomes `primary_phone`. This is **heuristic and error-prone** (~15%+
  of mined numbers are misattributed, e.g. a quoted signature credited to the
  wrong person), so it is off by default. Pass `--self-phone` to exclude your
  own number, and note that `--reconcile` drops any number shared across many
  contacts. Phones from vCard/Outlook/LinkedIn imports are always kept.
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

An import **always folds into the database -- it never wipes it.** A missing
database file is created fresh; an existing one is read, merged into, and written
back. There is no separate "overwrite the whole file" mode: to start over, delete
the file (or point `--db` at a new path).

For an existing contact, the counts (`num_emails`, `num_sent`, `num_received`)
are overwritten with the latest import, the email list is unioned, the
interaction date range widens, and `import_date` updates. Hand-edited text fields
(`friend`/`title`/`address`, plus name/company/phone) are **preserved** by
default -- the import only fills a blank. Pass **`--force`** to let the incoming
non-empty values **overwrite** those fields instead. Contacts present only in the
old file are always kept.

```bash
mc import archive.pst --db data/contacts.json          # fold in; keep manual edits
mc import archive.pst --db data/contacts.json --force  # let the import win on conflicts
mc import extra.json  --db data/contacts.json          # fold one DB into another
```

This lets you re-run as a mailbox grows, or accumulate multiple sources, without
losing manual annotations. Imported rows include blank columns for you to fill in
by hand: `friend`, `title`, and `address`. (The `category` column -- an industry
segment -- is filled automatically by `--reconcile` from the bundled yellowpages
directory.)

(Records are matched by email, or by LinkedIn profile URL when there is no email,
so email-less LinkedIn contacts survive a merge.)

### Importing from LinkedIn

A LinkedIn **Connections** export (`Settings -> Data privacy -> Get a copy of
your data -> Connections`) is the authority on a contact's *current* employer and
title. Import it with `--format linkedin` (the `.csv` extension alone is
ambiguous, so the format is explicit, like `outlook`); it folds into your existing
DB:

```bash
mc import Connections.csv --format linkedin --db data/contacts.json
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

### Scrape a mailbox for an LLM

`mc scrape` writes a per-email **JSONL** corpus (one JSON object per line) for
feeding to an LLM. The full body is included, HTML stripped to text. Every
message is included except obvious `no-reply` senders, and it works on both mbox
and PST.

```bash
mc scrape mailbox.mbox -o emails.jsonl
mc scrape archive.pst  -o emails.jsonl                      # works on PST too
mc scrape mailbox.mbox -o emails.jsonl --db contacts.json   # link to contacts
```

Pass `--db` (or set `$MC_DB`) to **link each email back to a contact**: the
record's `contact_id` is the `id` of the database contact the message is from or
to (the first matching address), or `null` if none matches. This couples the
corpus to your contacts DB -- you can group, filter, or attribute scraped email
by contact. Each record:

```json
{
  "message_id": "msg_987654321",
  "contact_id": "11111111-1111-1111-1111-111111111111",
  "from": "Jordan Vale <jordan@example.com>",
  "to": "you@example.org",
  "subject": "Follow up on hardware architecture",
  "date": "2025-03-15T14:22:00",
  "body": "Hey, great catching up at the conference..."
}
```

Without `--db`/`$MC_DB`, `contact_id` is `null` (the corpus is still produced).

The JSONL is streamed as messages are read, so it scales to very large mailboxes
without holding everything in memory.

## Exporting contacts

`mc export DEST` selects a subset of contacts from the database by per-column
criteria and writes the **whole record** for each match. The output format
follows the `DEST` extension:

- **`.csv`** -- all database columns. Pass `--format outlook` to instead write
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
# All semiconductor and defense contacts (category set by reconcile) -> vCard:
mc export leads.vcf --category "Semiconductor Devices,Defense"

# Everyone at Intel or AMD with at least 5 emails, active since 2024 -> CSV:
mc export intel_amd.csv --company Intel,AMD --min-emails 5 --last-after 2024-01-01

# All intel.com contacts since 2025 -> vCard:
mc export intel.vcf --email-domain intel.com --last-after 2025-01-01

# Only contacts at target-company domains, split across category files -> XLSX:
mc export targets.xlsx --whitelist semiconductor.txt defense.txt equipment.txt
```

The vCard maps name/emails (primary marked `PREF`), `company`->ORG,
`title`->TITLE, `phone_numbers`->TEL (primary marked `PREF`), `birthday`->BDAY,
`github`->URL, `address`->ADR/LABEL, and `category`/`friend`->CATEGORIES, plus a
NOTE with the email counts, last-contact date, and `notes`.

`--category` matches the industry segment assigned by `--reconcile` (e.g.
`Semiconductor Devices`, `Defense`, `Venture Capital`, `Academic`). See `mc -h`
for the full set of filters (`--category`, `--first-name`, `--last-name`,
`--email-domain`, `--whitelist`, `--blacklist`, `--min/max-emails`,
`--min/max-sent`, `--min/max-received`, `--first-after/before`,
`--last-after/before`).

## Reconciling contacts

After building a database from several sources (mbox, PST, vCard, Outlook,
LinkedIn), `mc reconcile` is a one-step cleanup pass over the database that merges
duplicates (by both name and email) and tidies records, in place:

```bash
mc reconcile --db contacts.json    # or just `mc reconcile` with $MC_DB set
```

In order, reconcile:

1. **Drops junk addresses** -- invalid emails, automated/bot senders, and
   role/generic mailboxes (`no-reply@`, `info@`, `sales@`, ...).
2. **Merges duplicates by shared email** -- the same person under two display
   names who share an address (e.g. `Bob Jones` and `Robert Jones`, both at
   `bob@acme.com`) that a name match misses. Free-provider (gmail/...), role, and
   bot addresses are **not** used as merge keys, so people who merely share a
   common mailbox are not fused.
3. **Merges duplicates by name** (same first+last, case-insensitive).
4. **Recomputes** `num_emails`, lowercases/dedupes `emails`, and fills a blank
   `company` from the primary domain.
5. **Picks the best primary email** -- the address whose domain matches the
   contact's current `company` if there is one, else the most-used address.
6. **Normalizes** name capitalization and phone numbers (to `+E.164`).
7. **Standardizes companies and assigns `category`** from the bundled
   yellowpages directory, keyed by email domain: company-name spelling variants
   are collapsed to one form, and listed domains get an authoritative company
   name plus an industry `category` (e.g. `Semiconductor Devices`, `Defense`); a
   `.edu` address gets `Academic`. Unlisted domains keep their company and a
   blank category.

When duplicates are merged, `company`/`title` come from the **LinkedIn-sourced**
record (LinkedIn is the authority on current employer); all other fields come
from the record with the **newest `last_interaction`**. Counts sum, emails and
sources union, dates widen. A record left with no email, LinkedIn URL, **or**
phone number is dropped. Reconcile is idempotent -- running it again changes
nothing.
