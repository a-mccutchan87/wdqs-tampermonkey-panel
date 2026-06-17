# WDQS Tampermonkey Panel

A Tampermonkey userscript panel for Wikidata Query Service that captures graph-style node results, deduplicates QIDs, exports JSON/CSV/XLSX, and shows acquired node relationships in a small tree summary viewer.

## Files

- `wdqs-capture-panel.user.js` — the Tampermonkey userscript behavior.
- `wdqs-panel.html` — the panel markup loaded as a Tampermonkey resource.
- `wdqs-panel.css` — the panel styles loaded as a Tampermonkey resource.

## Install

1. Keep this repository public so Tampermonkey can read the raw resource files.
2. Open `wdqs-capture-panel.user.js` in GitHub.
3. Click **Raw**.
4. Tampermonkey should prompt you to install or update the script.

## Notes

The userscript reads the HTML and CSS files through Tampermonkey `@resource` lines. If the repository name, branch, or username changes, update the raw GitHub URLs in the userscript header.
