# Verification of this source snapshot

2026-10-29, Windows, Python 3.12. Four offline unittest checks passed, including a complete synthetic Excel report roundtrip and unchanged input hash, both percentage formats, no-active-campaign rejection, and environment-only credential selection with all key-file reads prohibited. ResourceWarning was treated as an error; the Excel header reader now closes its workbook explicitly. Both Python files compile.

The WB script no longer contains the private shared credential-file path. The public copy loads only WB_TOKEN from the environment and refuses an empty value. Its network API/search functions were not executed or validated. No real report, seller credentials, API call, campaign update, GUI or browser was used. Financial heuristics were preserved, not evaluated against real campaigns. Only the new source copy was edited.
