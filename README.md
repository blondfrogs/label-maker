# Free Label Maker for Addresses

A completely free, web-based tool to generate address labels for Avery and compatible label sheets. No installation required - works entirely in your browser!

## Try It Now

🔗 **[Open Label Maker](https://blondfrogs.github.io/label-maker/)**

## Features

- 📋 **Multiple Label Templates** - Support for 6+ popular Avery formats
  - Avery 8160/5160 - Standard Address (30 per sheet) ⭐ Most Popular
  - Avery 5162/8162 - Wide Address (14 per sheet)
  - Avery 5167/8167 - Return Address (80 per sheet!)
  - Avery 5161, 5163, 5164 - Various sizes
- 📄 **Two Input Methods**
  - Upload CSV/TSV files (drag & drop)
  - Manual entry directly in the browser
- 💾 **PDF Output** - Perfect for Google Docs, Word, and all platforms
- 🔒 **100% Client-Side** - Your data never leaves your browser
- ✨ **No Installation** - No sign-up, no downloads, completely free
- 🏪 **Works with Store Brands** - Compatible with Office Depot, Staples, and generic labels

## How to Use

### Method 1: Upload a File
1. Select your label template from the dropdown
2. Drag and drop your CSV/TSV file (or click to browse)
3. Download the generated PDF
4. Print at 100% scale

### Method 2: Manual Entry
1. Select your label template
2. Click "Manual Entry" tab
3. Type addresses one at a time
4. Generate PDF when ready

## File Format

Your CSV or TSV file should have this format:

```csv
Fullname,Address,City State Zip Code
John Smith,123 Main St,New York NY 10001
Jane Doe,456 Oak Ave,Boston MA 02101
```

Or with separate columns:

```csv
Fullname,Address,City,State,Zip
John Smith,123 Main St,New York,NY,10001
Jane Doe,456 Oak Ave,Boston,MA,02101
```

## About Avery 8160

Avery 8160 labels are standard mailing labels with these specifications:
- **Label size:** 1" × 2⅝"
- **Labels per sheet:** 30 (3 columns × 10 rows)
- **Sheet size:** 8.5" × 11" (standard letter)

## Technical Details

Built with:
- Pure HTML, CSS, and JavaScript
- [jsPDF](https://github.com/parallax/jsPDF) library for PDF generation
- Exact Avery 8160 specifications (0.5" top margin, 0.19" side margins, 2.75" horizontal pitch)
- No backend required - fully static

## Local Development

Simply open `index.html` in any modern web browser. No build process or dependencies to install!

## Printing Instructions

**Important:** For perfect alignment:
- Print at **100% scale** (Actual Size) - NOT "Fit to Page"
- Use **US Letter** (8.5" × 11") paper size
- **Test first** - Print one page on regular paper and hold it up to your label sheet to verify alignment
- Do not override margins in printer settings

## License

MIT License - Completely free to use, modify, and distribute! See [LICENSE](LICENSE) file for details.

This tool is provided as-is with no warranty. Use at your own risk.

## Contributing

Issues and pull requests are welcome! This is an open-source project maintained by the community.

## Credits

- Built with [jsPDF](https://github.com/parallax/jsPDF) for PDF generation
- Avery is a registered trademark of Avery Dennison Corporation
- This is an independent tool not affiliated with Avery
