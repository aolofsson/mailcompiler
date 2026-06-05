# mailcompiler

Build and query a personal contacts database from a Gmail Takeout mailbox.
`mc import` turns an `.mbox` into a `contacts.csv`; `mc list` queries that
CSV and emits matching email addresses.

## Install (development)

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"          # provides the `mc` command
```

`mc` has one subcommand per operation, all with a uniform `-i/--input`:

```
mc import  -i MBOX -o OUT [...]               build/merge the contacts CSV
mc list    -i CONTACTS.csv [filters...]       list matching addresses
```

Without installing, run it as a module: `python -m mailcompiler.mailcompiler <command> ...`.

## Building the contacts database

`mc import` parses a Gmail Takeout `.mbox` into `contacts.csv`. Contacts are
people you have corresponded with (sent to or heard from), with automated/bulk
senders, spam, and nameless entries filtered out. Identities are merged by
display name, and company is derived from the email domain.

```bash
mc import \
  -i "/path/to/Takeout/Mail/All mail Including Spam and Trash.mbox" \
  -o data
```

`-i` and `-o` are required. `-o` may be a directory (writes `contacts.csv`
inside it) or a file path. See `mc import -h` for all options.

### What gets imported

A row is created for each **person you have corresponded with** -- anyone you
sent mail to or who sent mail to you. Specifically, an address is imported only
if **all** of these hold:

- It is a recipient (`To`/`Cc`) of mail you sent, **or** the sender (`From`) of
  mail you received (either direction qualifies).
- The message is **not** in Spam (Spam-labeled mail is skipped).
- It is **not one of your own addresses** (auto-detected from the `Delivered-To`
  header and the `From` of `Sent`-labeled mail).
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
- `company` is derived from the email domain (blank for free providers like
  gmail/yahoo/outlook), and each row records sent/received counts, the first and
  last interaction dates, and the `source` mbox filename.

Pass `--blacklist PATH` to exclude whole domains from the contacts. The file
lists one domain per line (`#` comments and blank lines ignored); entries may be
written as `example.com` or `@example.com`, and subdomains are matched too:

```text
# blacklist.txt
recruiting-spam.com
@newsletters.example.org
```

### Merge vs overwrite

If the output CSV already exists, the importer **merges** into it. For an existing
contact, the counts (`num_emails`, `num_sent`, `num_received`) are overwritten
with the latest import, the email list is unioned, and the interaction date range
widens; hand-edited fields (notably the `vip` column, plus name/company) are
preserved. Contacts present only in the old file are left untouched. This lets you
re-run as a mailbox grows, or accumulate multiple mailboxes, without losing manual
annotations.

Use `-f` / `--force` to ignore the existing file and write a fresh CSV instead
(this discards any manual edits such as `vip`).

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
mc list -i data/contacts.csv --vip

# Everyone at Intel or AMD with at least 5 emails, active since 2024:
mc list -i data/contacts.csv \
  --company Intel,AMD --min-emails 5 --last-after 2024-01-01

# All intel.com addresses you've corresponded with since 2025:
mc list -i data/contacts.csv \
  --email-domain intel.com --last-after 2025-01-01

# Every known address (not just primary) for a company, written to a file:
mc list -i data/contacts.csv \
  --company globex --all-emails -o segment.txt
```

See `mc list -h` for the full set of filters (`--vip`, `--first-name`,
`--last-name`, `--min/max-emails`, `--min/max-sent`, `--min/max-received`,
`--first-after/before`, `--last-after/before`).
