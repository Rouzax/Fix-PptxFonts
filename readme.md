# Fix-PptxFonts

PowerShell helper to **normalize fonts in an unpacked PowerPoint (.pptx)** by replacing hard-coded families with **theme font tokens**:

* `Arial Nova Light` (and variants) → `+mj-lt` (**Headings / Major Latin** → e.g., *Roboto Light*)
* `Arial` → `+mn-lt` (**Body / Minor Latin** → e.g., *Roboto*)

This lets slides inherit fonts from your theme’s `<a:fontScheme>` instead of being stuck with Arial overrides.

---

## Why?

Templates often contain shapes with explicit `typeface="Arial"` in the XML, which **overrides** your theme. This script scrubs those overrides so your theme (Major/Minor) takes control everywhere—titles use your heading font, body text uses your body font.

---

## What it does

* Recurses your **unpacked** `.pptx` folder (only under `ppt\`).
* **Excludes** `ppt\theme\` by default (so your theme files are not changed). You can opt in with `-IncludeTheme`.
* Rewrites any `typeface="..."` attributes in `.xml` and `.rels` files:

  * `Arial Nova Light` (common variants) → `+mj-lt`
  * `Arial` → `+mn-lt`
* Optional: with `-IncludeVariants`, also maps `ArialMT`, `ArialPSMT`, `Arial Unicode MS`, `Arial Narrow`, `Arial Black` → `+mn-lt`.
* Writes files as **UTF-8 without BOM**.
* Creates `.bak` backups by default.

> It does **not** change your theme unless you explicitly pass `-IncludeTheme`.

---

## Requirements

* Windows PowerShell **5.1** or PowerShell **7+**
* An **unpacked** `.pptx` (it’s a ZIP):

  * Rename `MyDeck.pptx` → `MyDeck.zip` and extract, **or**
  * Use any zip tool to extract it.
* Your root folder should contain a `ppt\` subfolder:

  ```
  MyDeck_unpacked/
  ├─ [Content_Types].xml
  ├─ _rels/
  └─ ppt/               ← the script focuses here
     ├─ slides/
     ├─ slideMasters/
     ├─ slideLayouts/
     └─ theme/          ← excluded by default
  ```

---

## Install

Copy `Fix-PptxFonts.ps1` into your repo (e.g., at the root).

---

## Usage

### Dry run (preview changes)

```powershell
.\Fix-PptxFonts.ps1 -Root "D:\PowerPoint\MyDeck_unpacked" -DryRun
```

### Apply changes (with backups)

```powershell
.\Fix-PptxFonts.ps1 -Root "D:\PowerPoint\MyDeck_unpacked"
```

### Also map common Arial variants to Body

```powershell
.\Fix-PptxFonts.ps1 -Root "D:\PowerPoint\MyDeck_unpacked" -IncludeVariants
```

### Include theme files (normally excluded)

```powershell
.\Fix-PptxFonts.ps1 -Root "D:\PowerPoint\MyDeck_unpacked" -IncludeTheme
```

### No backups

```powershell
.\Fix-PptxFonts.ps1 -Root "D:\PowerPoint\MyDeck_unpacked" -NoBackup
```

**Parameters**

* `-Root` *(required)*: Path to the folder that contains `ppt\`.
* `-DryRun`: Report matches without modifying files.
* `-NoBackup`: Don’t create `.bak` backups before writing.
* `-IncludeVariants`: Remap common Arial variants to `+mn-lt`.
* `-IncludeTheme`: Include `ppt\theme\` files in processing.

---

## Sample output

```
D:\PowerPoint\MyDeck_unpacked\ppt\slideMasters\slideMaster2.xml
  Headings (Arial Nova Light -> +mj-lt): 3
  Body (Arial -> +mn-lt): 22
D:\PowerPoint\MyDeck_unpacked\ppt\slides\slide1.xml
  Headings (Arial Nova Light -> +mj-lt): 0
  Body (Arial -> +mn-lt): 5
DONE. Files changed: 2 | Headings replacements: 3 | Body replacements: 27
```

---

## Repack the deck

1. Zip the **contents** of the unpacked folder (not the folder itself).
2. Rename the `.zip` back to `.pptx`.
3. Open in PowerPoint.
4. **Reset** slides (Home → Reset) if any placeholders were manually formatted.
5. Optionally use **Home → Replace → Replace Fonts…** to mop up any leftover oddities.

---

## Important notes

* **Theme mapping:**
  `+mj-lt` = Major/Latin (Headings), `+mn-lt` = Minor/Latin (Body). Configure these in your theme, e.g.:

  ```xml
  <a:fontScheme name="Font 2026">
    <a:majorFont>
      <a:latin typeface="Roboto Light"/>
      <a:ea typeface="Noto Sans JP"/>
      <a:cs typeface="Noto Sans Arabic"/>
    </a:majorFont>
    <a:minorFont>
      <a:latin typeface="Roboto"/>
      <a:ea typeface="Noto Sans JP"/>
      <a:cs typeface="Noto Sans Arabic"/>
    </a:minorFont>
  </a:fontScheme>
  ```
* **Placeholders vs shapes:**
  Placeholders (`<p:ph type="title"/>`, `<p:ph type="body"/>`) inherit from the theme. Plain shapes with direct font runs can override. This script removes those overrides.
* **Encoding:**

  * On **PS7+**, `Set-Content -Encoding utf8` is BOM-less.
  * On **PS5.1**, the script uses .NET to write **UTF-8 without BOM**.
* **Safety:**
  `.bak` backups are created by default. Use `-NoBackup` to disable.

---

## Extending mappings

Want to map more fonts (e.g., Calibri → Body)? Add to the variants regex or create a new one:

```powershell
# Example: also map Calibri to body
$reCalibri = '(?i)(typeface\s*=\s*["''])Calibri(["''])'
$text = [regex]::Replace($text, $reCalibri, '$1+mn-lt$2')
```

Or broaden the “major” regex if you find other `Arial Nova Light` spellings:

```powershell
$reMajor = '(?i)(typeface\s*=\s*["''])(?:Arial\s*Nova\s*Light|ArialNova[-\s]?Light|Arial\s*Nova\s*Lt|ArialNovaLt)(["''])'
```

---

## Troubleshooting

* **“UTF8NoBOM is invalid” on PS5.1:** The script already handles this by using .NET for BOM-less UTF-8.
* **No matches for “Headings”:** Your files may use a different name for Arial Nova Light. Use the broadened `$reMajor` above.
* **Weird characters like `â†’`** in console output: that’s Unicode punctuation from older scripts; this repo uses ASCII arrows `->`.

---

## Contributing

PRs welcome! Please:

* Keep the script **ASCII-only** (no smart quotes/dashes).
* Support both Windows PowerShell 5.1 and PowerShell 7+.
* Add tests or sample XMLs where helpful.

---

## License

MIT. Use at your own risk.
