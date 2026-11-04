# factory-invoice-to-opdn

Agent skill: factory Invoice + Packing List (.xls) -> SAP B1 GRPO upload workbook (OPDN_Header / PDN1_Line).

Layout
- SKILL.md                         agent instructions (under 8,000 chars)
- references/                      anatomy, field mapping, column defaults (CSV), validation rules
- scripts/build_opdn.py            reference implementation; python scripts/build_opdn.py <factory.xls> --lookup <po-lookup-table.xlsx>
- scripts/xlrd/                    bundled xlrd 2.0.1 (reads .xls) for sandboxes without it
- scripts/make_column_defaults.py  regenerates references/opdn-column-defaults.csv from the template
- examples/                        template, per-invoice output for 4762/4763/4764, *_per-po.xlsx variant, *_with-lookup.xlsx (joined against po_lookup_SAMPLE_fake_values.csv)

Requires: python3, openpyxl (xlrd bundled).
