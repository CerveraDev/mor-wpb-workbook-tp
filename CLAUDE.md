# Project Notes

## Google Sheets IMAGE() Formula

To embed slide thumbnails in Google Sheets, use:

```
=IMAGE("https://raw.githubusercontent.com/CerveraDev/mor-wpb-workbook-tp/<branch>/slide-thumbnails/slide-XX.jpg")
```

- Replace `<branch>` with the active branch name (e.g. `claude/sales-talking-points-90REn`)
- Replace `XX` with the zero-padded slide number (e.g. `01`, `12`, `43`)
- The formula is placed in the Thumbnail column (column B in the workbook)
- Thumbnails are sized 100×56px (16:9 aspect ratio)
- The same formula is generated programmatically in `generate_xlsx.py` via `BASE_URL`
