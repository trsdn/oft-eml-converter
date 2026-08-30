# Decision records

One decision per file. A record states what was chosen, what it cost, and what would have to change for the decision to be revisited.

Records are immutable once merged. A decision that no longer holds is superseded by a new record that names it, rather than edited in place — otherwise the reasoning behind the current state disappears and the same argument gets relitigated every year.

| Record | Decision |
|---|---|
| [0001](0001-delegate-msg-parsing.md) | MSG and OFT parsing is delegated to `extract-msg` |
| [0002](0002-separate-core-from-gui.md) | The conversion core is independent of the GUI |
