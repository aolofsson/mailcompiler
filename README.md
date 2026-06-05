# mailcompiler

Build and query a personal contacts database from a Gmail Takeout mailbox.
`mc import` turns an `.mbox` into a `contacts.json` database; `mc list` queries
it and emits matching email addresses.

## Install (development)

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"          # provides the `mc` command
```

`mc` has one subcommand per operation, all with a uniform `-i/--input`:

```
mc import  -i MBOX -o OUT [...]               build/merge the contacts DB
mc list    -i CONTACTS [filters...]           list matching addresses
```

The database is JSON (the native format). Pass `-o something.csv` to export
CSV instead.

Without installing, run it as a module: `python -m mailcompiler.mailcompiler <command> ...`.

## Building the contacts database

`mc import` parses a Gmail Takeout `.mbox` **or** an Outlook `.pst` (chosen by
the `-i` file extension) into `contacts.json`. Contacts are people you have
corresponded with (sent to or heard from), with automated/bulk senders, spam,
and nameless entries filtered out. Identities are merged by display name, and
company is derived from the email domain.

```bash
mc import -i "/path/to/Takeout/Mail/All mail Including Spam and Trash.mbox" -o data
mc import -i "/path/to/archive.pst" -o data       # Outlook PST
```

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
(notably the `vip` field, plus name/company) are preserved. Contacts present only
in the old file are left untouched. This lets you re-run as a mailbox grows, or
accumulate multiple mailboxes, without losing manual annotations.

Use `-f` / `--force` to ignore the existing file and write a fresh one instead
(this discards any manual edits such as `vip`).

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

## Querying contacts

`mc list` selects a subset of contacts by per-column criteria and prints
their email addresses on one comma-separated line, ready to paste into a mail
client To: field. Text filters take comma-separated lists (case-insensitive,
match any); numeric and date filters are inclusive ranges; all filters combine
with AND.

Note: `--company` matches the derived company *name* (e.g. `Globex`), while
`--email-domain` matches the address domain (e.g. `globex.com`).

```bash
# Only contacts you have flagged VIP (non-empty vip column):
mc list -i data/contacts.json --vip

# Everyone at Intel or AMD with at least 5 emails, active since 2024:
mc list -i data/contacts.json \
  --company Intel,AMD --min-emails 5 --last-after 2024-01-01

# All intel.com addresses you've corresponded with since 2025:
mc list -i data/contacts.json \
  --email-domain intel.com --last-after 2025-01-01

# Every known address (not just primary) for a company, written to a file:
mc list -i data/contacts.json \
  --company globex --all-emails -o segment.txt
```

See `mc list -h` for the full set of filters (`--vip`, `--first-name`,
`--last-name`, `--min/max-emails`, `--min/max-sent`, `--min/max-received`,
`--first-after/before`, `--last-after/before`).
