# Manifest and mapping contract

## Manifest

Create `run.json` in or near the selected run folder. This is a concrete example, not permission to search for matching files:

```json
{
  "run_id": "client-run-1",
  "run_folder": ".",
  "company": "Example Co",
  "site_hosts": ["www.example.com"],
  "asset_hosts": ["cdn.shopify.com"],
  "output_date": "2026-10-15",
  "sources": [
    {
      "file": "11.1-11.9_Example Co_sitemap_audit_2026-10-13.xlsx",
      "topic": "sitemap_audit",
      "date": "2026-10-13"
    }
  ]
}
```

`run_folder`, optional `template`, and optional `profile` resolve against the manifest's directory. The default template/profile are bundled with the skill. Each source `file` must be a relative path inside the run folder (including resolved symlinks). Only listed files are read. Output goes in `run_folder`. Require exact lowercase hostnames; add www/non-www variants only when the run scope supports both. `asset_hosts` applies to Destination resource URLs only. A report without structured URLs produces an identity-verification warning; a filename alone does not independently establish its host.

`date` is the source evidence/filename date. `output_date` is an explicit report/consolidation date; neither silently overrides the other. Mixed evidence dates are allowed only because the exact reports belong to the explicitly selected run, and remain visible in provenance. No newest-file selection.

Old filenames lacking numeric prefixes need per-source `"legacy_name": true`. Company/topic/date are still checked. A modern prefix must match the actual scope even with legacy permission. A partial source report needs explicit `"checks": ["17.2"]`; otherwise all profile IDs for that topic are required. Multiple disjoint subset reports are supported; overlapping IDs or combined/split candidates stop.

## Verified source families

| Topic key | Expected IDs | Existing combined detail sheet |
|---|---|---|
| robots_audit | 10.1–10.7 | 10. Robot.txt / 10. Robots.txt |
| sitemap_audit | 11.1–11.9 | 11. Sitemap |
| pagination_audit | 14.1–14.6 | 14. Pagination |
| hreflang_audit | 15.1–15.3 | 15. HREFLANG |
| robots_meta_tags_audit | 12.1–12.3 | 12. Robots meta tags (conditional) |
| page_titles_audit | 17.1–17.4 | 17. Page Title |
| meta_descriptions_audit | 18.1–18.3 | 18. meta_description |
| headings_audit | 19.1–19.3 | 19. Heading |
| images_audit | 20.1–20.2 | 20. Image |

The current HTTPS producer uses Checklist and 9. Implement HTTPS Protocol. Its detail headers are Issue Type, Address, Related URL, Findings and Instructions. Issue Type begins with 9.1, 9.2 or 9.3. Address is a site URL; optional Related URL is a resource URL for 9.2 and a site URL for 9.1/9.3. An empty table is supported. Unsupported combined statuses remain an explicit error. GSC is deliberately absent. Pagination/Hreflang filenames accept both `onsite_pagination_audit` / `onsite_hreflang_audit` and their shorter registered tokens.

All observed Checklists are header row 1 with `Item No.`, `Item Name`, `Initial Check` or `Check Result`, `Findings`, `Instruction`. Robots instead embeds `10.x` in `Check`, with `Result`, `Findings`, `Instruction`. Headers match registered aliases case-insensitively after trimming/collapsing spaces; column order may vary. Duplicate matching aliases stop.

Issue tables use header row 1. Read IDs from Item No., registered embedded-ID columns, validated child-ID sheet prefixes or exact registered sheet aliases. All available IDs must agree. Sitemap uses a leading ID in Sitemaps, whose full value is a finding label. URLs Example holds actual URL evidence: 11.7 X uses one URL per row; other sitemap checks may use newline-separated URLs or N/A where the producer permits no affected URL. Preserve original text and row counts. Only single-URL cells get native URL hyperlinks.

Current content reports use X-only split sheets and a space after the filename range. Register 18. Descriptions - Missing as 18.1. Full Checklist coverage is required; non-X split detail tabs may be absent, but every X needs nonempty detail evidence. Retain support for existing combined layouts and underscore filenames. Reject overlapping tables, conflicting IDs and unmapped sheets. Optional-table behavior is limited to registered content layouts and the pagination, hreflang and robots meta tags contracts documented here; all other topics retain their existing requirements. Record omitted tables distinctly from observed empty tables in provenance.

## Destination contract

`scripts/profiles.json` owns exact destination worksheet names. Existing template row positions are discovered from B check IDs, not hardcoded; floating point artifacts such as `11.299999999` normalize only for numeric IDs. Text IDs remain exact (`11.10` never becomes `11.1`). D/E/F must be writable non-formula, non-hyperlink cells. G and populated detail A1 are controlled navigation cells: rebuild their links, replacing stale HYPERLINK formulas only in these cells.

`main_sheet` identifies the input template sheet `SEO Implementation Checklist`; `output_main_sheet` is `(auto) SEO Checklist`, explicitly requested by the user. The full prefixed original title would exceed Excel's 31-character limit. Rename only the generated copy and update internal hyperlinks, formulas, named references and chart series references. Source workbooks and the bundled template retain their original sheet names.

The output status vocabulary is `✔ / ✖ / NA / - / Human check`. Source `Human Check` normalizes only its capitalization. Manual checks retain their reason and unavailable issue count. The status dropdown is extended for onsite cells; series 7 keeps its original dropdown and values. Missing reports use `-` plus `Not checked`, not Human check.

Detail headers are row 3, except Robots row 4. Existing notes below headers are preserved; imports start below the last occupied row with a gap. Match existing headings to exact source headings first, then registered semantic aliases. Fill only these fields. Do not append columns, summary blocks or explanatory rows. Original source values, provenance and output row/cell mappings remain in the validation sidecar, including unmapped fields. Stop if an entire detail row maps to no existing field. Main G targets the first populated detail cell for its check, or the existing header for a check with no details. Detail A1 returns to the first corresponding Checklist B row. Exact targets are stored in navigation_links and checked after saving and during verify.

A split destination already exists in the supplied template for each 17/18/19/20 child check. Source naming changes therefore need no destination renaming. Preserve precise native names such as `18. Descriptions - Missing` for 18.1.

## Agent adaptation

For new names/headers, inspect only the explicit files. Copy the profile into the run folder and set `profile` in the manifest. Update source sheet aliases, filename tokens, field aliases or exact targets only when the IDs and column meaning establish the mapping. Never rename one metric to another (for example Character Count to Pixel Width). Extra source fields are preserved in the sidecar, not added to the destination schema or silently interpreted.

Different header rows, non-tabular issue layouts, merged input data, formula-derived results and new status meanings need a tested reader adapter; JSON aliases alone cannot establish their semantics. Unknown result tokens stop. Source formula cells stop; the tool does not execute workbook functions or rely on stale cached results.

Input results remain authoritative. This consolidator does not adjudicate SEO rules or improve findings. Missing reports, manual checks and incomplete evidence remain visible. Counts are observed rows, not a claim that all defects have been found; preserve source units and count limits.

## Current verification limits

The real sample covers eight onsite reports, 37 check results and 1,086 problem-detail rows. GSC is excluded. The template has known inherited broken Summary formulas and Google-specific formulas/anchors. The script checks preservation, requested writes, URL links and workbook navigation; it does not recalculate those original formulas or promise their Excel compatibility.


Pagination and hreflang always require complete selected Checklist coverage. Their registered producers create 14. Pagination or 15. HREFLANG only when at least one result is X. A Checklist-only report with no X is valid; preserve pass, N/A, Human check and other supported non-X statuses and all evidence limitations without reclassifying them. Record absent tables as omitted by producer contract, separately from observed header-only tables. Every X still requires nonempty correctly identified detail evidence. An absent or empty X table and missing Checklist rows remain errors. Existing registered combined-sheet aliases and populated-table validation remain unchanged. Zero recorded rows is not a confirmed issue count for Human check, N/A or unknown results.

## Robots meta tags producer contract

Use manifest topic robots_meta_tags_audit for all three IDs 12.1–12.3. Filename tokens onsite_robots_meta_tags_audit and robots_meta_tags_audit are explicitly registered; validate company, date and full range as usual. Checklist has Item No., Item Name, Check Result, Findings and Instruction. Require all three rows and nonblank names and narratives. Normalize existing status aliases to ✔, ✖ or Human check and retain raw spelling; 12.2 cannot be ✖. Unknown or incompatible results stop without reclassification.

The combined source sheet 12. Robots meta tags is optional when there are no failures. If supplied, require Issue, URLs, How to Fix and one Status / Findings (reviewed YYYY-MM-DD) header whose valid date equals the source report date. Each populated row starts Issue with its check ID, supplies nonblank issue text and fix, has one valid site URL, and uses Pending implementation in the dated status cell. Only failed 12.1 or 12.3 checks may have rows. Every failed check requires its own evidence. Split sheets, missing headers, malformed or foreign URLs, conflicting IDs and unsupported statuses stop. A valid empty non-failure table is observed, not omitted. Keep source_schema robots-meta-tags-conditional-v1 in both cases.

Discover destination Checklist rows by B IDs. In the pinned original template, 12.1–12.3 occupy rows 77–79; numeric floating-point 12.3 is normalized by the existing check-ID reader. Use existing 12. Robots meta tags: URLs maps to Address and Issue to Issues. No native field exists there for How to Fix or the dated status header; preserve their complete values in validation JSON, while Checklist F preserves the source check-level Instruction. Do not append fields or fabricate detail rows. Existing screenshots, notes, formulas and template limitations remain. Zero-row checks link to A3; populated checks link to their first mapped detail cell; A1 returns to Checklist B77 in this template. Verify discovered targets on other supported templates rather than hardcoding those coordinates.

Preserve English source narratives and all original evidence, including Chinese source text, source report dates and historical crawl caveats. Generate consolidation annotations in English. Omission is not proof of zero defects: Human check retains an unconfirmed count. This adapter neither reruns audits nor adds GSC or canonical-route coverage.
