# 0001 — Delegate MSG and OFT parsing to extract-msg

- Status: accepted
- Date: 2025-09-05

## Context

`.oft` files are Outlook templates in the Compound File Binary Format, carrying MAPI properties. The interesting content — sender, recipients, subject, HTML body, plain-text alternative, attachments, and the Content-ID mapping that makes inline images work — is spread across property streams whose encoding depends on how Outlook wrote the file.

Parsing this by hand is possible. It is also a long-term commitment to tracking a format nobody documents in full, on behalf of a tool whose actual job is producing a MIME message.

## Decision

All binary parsing is delegated to [`extract-msg`](https://github.com/TeamMsgExtractor/msg-extractor). This repository owns the mapping from parsed properties to a `.eml` file and nothing below that line.

## Consequences

What this buys:

- Format edge cases are somebody else's maintained problem, tested against a far wider corpus of real files than this project will ever see.
- The conversion core stays small enough to read in one sitting.

What it costs:

- A runtime dependency on a library that is licensed GPL-3.0 while this repository is MIT. The dependency is installed separately via pip and is not vendored or redistributed, so the licences do not conflict, but this is the reason the boundary must stay a dependency boundary.
- Upstream bugs are not directly fixable here.

## Revisit when

`extract-msg` becomes unmaintained, or a required capability is rejected upstream. Neither has happened. Writing a parser because a specific file failed is not a reason — that is a bug report.

## Note for agents

This decision is the reason `AGENTS.md` forbids hand-written MSG parsing. An agent asked to "fix OFT parsing" should reach for the upstream library, not for `struct.unpack`.
