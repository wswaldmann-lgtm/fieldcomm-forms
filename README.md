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
| FC-FORM-006 | Photo Log | yes | yes | MOD-02 Photo Log |
| FC-FORM-007 | 3-Week Look-Ahead | yes | yes | MOD-03 Schedule |
| FC-FORM-008 | Delivery & Material Receiving Log | yes | yes | MOD-05 Materials & Delivery |
| FC-FORM-009 | Credential & Scope Authorization | yes | | MOD-07 Credential Gate (FC-SPEC-003) |
| FC-FORM-010 | Incident & Near-Miss Report | yes | | MOD-11 Incident & Near-Miss |
| FC-FORM-011 | Work Completion Attestation | yes | | MOD-16 Attestation + Draw Release |
| FC-FORM-012 | COI & Lien Waiver Tracker | yes | yes | MOD-18 COI + Lien Waiver |
| FC-FORM-013 | Closeout Checklist | yes | | MOD-24 Closeout Package |

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
