# 0002 — Keep the conversion core independent of the GUI

- Status: accepted
- Date: 2025-09-05

## Context

The project ships two entry points: a command-line script and a Tkinter GUI. The obvious shortcut is to let them share state, or to let the converter report progress by touching widgets directly.

Tkinter is also the least portable part of the stack. It is absent from some Python builds, needs a system package on Linux, and cannot be imported at all on a headless CI runner without a display.

## Decision

`oft_to_eml_converter.py` holds all conversion logic and imports nothing from the GUI. `oft_to_eml_gui.py` imports the core. The dependency points one way and never the other.

## Consequences

What this buys:

- The core is testable on a headless runner, which is what makes the cross-platform CI matrix possible at all.
- A missing `tkinter` degrades the tool to CLI-only instead of breaking it entirely.
- Progress reporting is plain stdout, which is also why the CLI output stays usable without colour or Unicode decoration.

What it costs:

- The GUI cannot push rich progress state through the core without an explicit callback seam. Today it does not need one.

## Revisit when

A third entry point needs structured progress events rather than printed lines. At that point the seam should be an explicit callback or an event object, not an import in the other direction.
