# Deal Memo Publisher

`Code.gs` is the script bound to the deal model Sheet (master template
`06_Multifamily_Deal_Template_v3.8_MASTER`; every deal copy inherits it). It adds the
**Deal Memo** menu: *Refresh numbers & publish*.

This file is the source of truth. If the script in a Sheet is ever lost, copy it from here:
Sheet > Extensions > Apps Script > replace everything > Save > reload the Sheet.

- `_shells/<Sheet ID>.html` — one memo shell per deal, written by Claude at "render this deal".
  It holds the written sections and the deal's settings; the script fills in the numbers.
- `<slug>.html` at the top level — the published memos.

Token math mirrors render spec v1.7. Change one only together with the other.
