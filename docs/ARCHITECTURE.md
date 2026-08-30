# Architecture

How the converter is put together, and the constraints that are not obvious from reading the source. For *why* the two central choices were made, see [decisions](decisions/).

## Shape

```text
┌──────────────────────┐        ┌──────────────────────┐
│  oft_to_eml_gui.py   │───────▶│ oft_to_eml_converter │
│  Tkinter, presenta-  │ import │ .py                  │
│  tion only           │        │ parsing + MIME       │
└──────────────────────┘        └──────────┬───────────┘
                                           │
                             ┌─────────────┴─────────────┐
                             │                           │
                      ┌──────▼──────┐            ┌───────▼───────┐
                      │ extract-msg │            │ email (stdlib)│
                      │ reads .oft  │            │ writes .eml   │
                      └─────────────┘            └───────────────┘
```

The dependency arrow points one way. The core never imports the GUI, which is what keeps it usable on a headless CI runner.

## Conversion pipeline

1. **Open** the `.oft` with `extract-msg`, which resolves the Compound File Binary container and its MAPI property streams.
2. **Read headers** — sender, recipients, CC, subject, date — and map them onto MIME headers.
3. **Collect bodies.** A message may carry a plain-text body, an HTML body, or both. Both are preserved when present.
4. **Partition attachments** into two groups. An attachment carrying a Content-ID is an inline image referenced from the HTML body; everything else is a regular attachment. Getting this wrong is what makes images appear as duplicated file attachments instead of rendering in place.
5. **Assemble the MIME tree.** Inline parts go into a `multipart/related` alongside the HTML body; the text and HTML alternatives go into `multipart/alternative`; regular attachments wrap the whole thing in `multipart/mixed`.
6. **Write** the `.eml` in UTF-8.

Step 4 and step 5 are the parts worth being careful with. Everything else is bookkeeping.

## Constraints

**No network at runtime.** The converter opens local files and writes local files. It never resolves a remote reference, not even one embedded in the HTML body it is copying. This is what allows the README to state that nothing leaves the machine, and it is a property to preserve rather than an accident.

**Plain ASCII output.** Progress and error messages avoid colour, Unicode box drawing, and emoji so the CLI stays readable on a terminal without colour support and through a screen reader. Windows console encoding has broken this project before; the tests cover it.

**Headless-safe imports.** `tkinter` is imported by the GUI module only, at module level, and the CI suite asserts the core imports cleanly without a display.

**The dependency surface is deliberately narrow.** One runtime dependency plus the standard library. Widening it needs a reason, because the tool's value is partly that it installs anywhere with minimal friction.

## Testing

`tests/test_converter.py` covers the conversion core, the CLI error paths, and the headless-import guarantee. The CI matrix runs the suite across Windows, macOS, and Linux on every supported Python version, because the failures this project actually sees are platform encoding failures rather than logic errors.
