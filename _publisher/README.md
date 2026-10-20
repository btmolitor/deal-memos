# Deal Memo Publisher

`Code.gs` is the script bound to the deal model Sheet (master template
`06_Multifamily_Deal_Template_v3.8_MASTER`; every deal copy inherits it). It adds the
**Deal Memo** menu: *Refresh numbers & publish*, and *Export Summary PDF to deal folder*
(dated landscape PDF of the Summary tab, saved next to the Sheet).

This file is the source of truth. If the script in a Sheet is ever lost, copy it from here:
Sheet > Extensions > Apps Script > replace everything > Save > reload the Sheet.

- `_shells/<Sheet ID>.html` — one memo shell per deal, written by Claude at "render this deal".
  It holds the written sections and the deal's settings; the script fills in the numbers.
- `<slug>.html` at the top level — the published memos.
- `_shells/comp_map_template.html` — the comp map shell (one token, `COMP_DATA`). Refresh fills it from the
  deal Sheet's Rent Comps tab and publishes `<slug>_comps.html` next to the memo; the memo's section 09 links to it.
  Blank Lat/Lng cells on Rent Comps are geocoded (Census Bureau, no key) and written back on each refresh.

Token math mirrors render spec v1.8. Change one only together with the other.
