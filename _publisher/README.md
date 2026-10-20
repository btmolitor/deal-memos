# Deal Memo Publisher

Two files:

- **`Loader.gs`** — the only code that lives inside a Sheet (master template
  `06_Multifamily_Deal_Template_v3.8_MASTER` and every deal copied from it). It builds the **Deal Memo**
  menu and, on every click, fetches the current `Code.gs` from this repo and runs it. Pasted once per Sheet;
  after that, changes to `Code.gs` are live everywhere with no paste. The menu items themselves are the one
  thing that still needs a paste when they change.
- **`Code.gs`** — the publisher: *Refresh numbers & publish* (memo + comp map + Deal Analyzer List),
  *Preview only*, *Export Summary PDF to deal folder*, *Open live memo*, *Check setup*, *Set GitHub key*.
  Edit this file in the repo; nothing else to do.

If the script in a Sheet is ever lost, or a Sheet still carries an old copy of `Code.gs`:
Sheet > Extensions > Apps Script > replace everything with **`Loader.gs`** > Save > reload the Sheet.

- `_shells/<Sheet ID>.html` — one memo shell per deal, written by Claude at "render this deal".
  It holds the written sections and the deal's settings; the script fills in the numbers.
- `<slug>.html` at the top level — the published memos.
- `_shells/comp_map_template.html` — the comp map shell (one token, `COMP_DATA`). Refresh fills it from the
  deal Sheet's Rent Comps tab and publishes `<slug>_comps.html` next to the memo; the memo's section 09 links
  to it and the Rent Comps caption row (above the `Property` header) gets the same link.
  Blank Lat/Lng cells on Rent Comps are geocoded (Census Bureau, no key) and written back on each refresh.

Token math mirrors render spec v1.8. Change one only together with the other.
