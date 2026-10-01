# FieldComm Forms

Free construction field forms: the clipboard paperwork that has run job sites
for decades, as clean fillable PDFs and Excel logs.

| Code | Form | PDF | Excel | FieldComm Pro module |
|---|---|---|---|---|
| FC-FORM-001 | Daily Report | yes | | MOD-01 Daily Report |
| FC-FORM-002 | RFI Request & Log | yes | log | MOD-14 RFI Manager |
| FC-FORM-003 | Punch List | yes | yes | MOD-15 Punch List |
| FC-FORM-004 | Change Order Request & Log | yes | log | MOD-19 Change Order Manager |
| FC-FORM-005 | Toolbox Talk Sign-In | yes | | MOD-10 Toolbox Talk |

The site (`index.html`) publishes to GitHub Pages on every push to `main`
(`.github/workflows/pages.yml`). One-time setup: Settings → Pages → Source:
**GitHub Actions**.

## Rebuild the forms

```
pip install reportlab openpyxl
python3 tools/build_forms.py
```

Every form is defined as data in `tools/build_forms.py`; add a form there and
re-run. Commit the regenerated files in `forms/`.

Templates only, not legal advice. AIA documents are not reproduced; required
government forms (OSHA logs, statutory lien waivers) should come from the
official source.
