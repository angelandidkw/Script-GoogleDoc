# Script-GoogleDoc
I have a break through in google doc. An accurate, automated Table of Contents generator for Google Docs that calculates precise page numbers using PDF export and weighted content analysis.

# Google Docs Table of Contents Generator

An accurate, automated Table of Contents generator for Google Docs that calculates precise page numbers using PDF export and weighted content analysis.

## Features

- ✅ **Accurate Page Numbers** - Uses PDF export metadata and weighted content calculation
- ✅ **Smart Content Weighting** - Considers font sizes, spacing, tables, images, and list items
- ✅ **Page Break Detection** - Respects explicit page breaks in your document
- ✅ **Automatic TOC Updates** - Regenerates TOC entries while preserving document structure
- ✅ **Professional Formatting** - Beautiful, consistent styling for Heading 1 and Heading 2

## Installation

1. Open your Google Doc
2. Go to **Extensions → Apps Script**
3. Delete any existing code
4. Paste the contents of `Code.gs`
5. Save the project (name it "TOC Generator")
6. Refresh your Google Doc
7. You'll see a new menu: **TOC Generator**

## Usage

1. Add a heading in your document that includes the text "Table of Contents"
2. Use **Heading 1** and **Heading 2** styles for your document sections
3. Click **TOC Generator → Update Table of Contents**
4. Your TOC will be generated with accurate page numbers!

## The Problem We Solved

### Original Issues

The original implementation had several accuracy problems:

1. **Simplistic Length Calculation**
   - Used `text.length + 50` for all content
   - Ignored that 18pt headings take more space than 11pt body text
   - Didn't account for paragraph spacing, line spacing, or margins

2. **Fixed Table Weights**
   - All tables counted as 500 units regardless of size
   - A 2x2 table was treated the same as a 50x10 table

3. **Missing Element Types**
   - Images weren't properly weighted
   - List items weren't considered
   - Page breaks were ignored

4. **Critical Spacing Bug**
   - Page numbers calculated AFTER including the heading's own weight
   - Each heading was positioned as if its content was already rendered
   - This pushed all page numbers too far forward

### The Solution

#### 1. Weighted Element Calculation (`calculateElementWeight`)

```javascript
// Before: text.length + 50
// After: Comprehensive weighting
```

Now considers:
- **Font size** - Larger fonts = more vertical space
- **Line spacing** - 1.5 line spacing vs single spacing
- **Paragraph spacing** - Space before and after
- **Estimated lines** - Multi-line paragraphs properly calculated
- **Table complexity** - Actual row count and cell content
- **List item overhead** - Indentation and bullet point space
- **Images** - Appropriate vertical space allocation

#### 2. Enhanced Page Number Detection (`getImprovedPageNumbers`)

- **Multiple PDF parsing methods** for reliable page count extraction
- **Page break detection** using `DocumentApp.Attribute.PAGE_BREAK_BEFORE`
- **Two-pass algorithm**:
  1. Calculate total document weight
  2. Assign proportional pages based on cumulative weight
- **Critical fix**: Calculate page number BEFORE adding element's own weight

#### 3. Accurate TOC Page Estimation (`estimateTOCPages`)

- Reads actual document dimensions (page height, margins)
- Calculates usable vertical space
- Estimates TOC entries based on font size + spacing
- Properly accounts for Heading 1 vs Heading 2 differences

## How It Works

### Algorithm Overview

```
1. Export document as PDF to get actual total pages
2. Calculate "weight" for every element in document
3. For each heading:
   a. Sum all content weight BEFORE the heading
   b. Calculate proportion: beforeWeight / totalWeight
   c. Multiply by total pages → heading's page number
   d. Add TOC page offset
   e. THEN add the heading's weight for next calculation
```

### Weight Calculation Formula

**Paragraphs:**
```
weight = (characters × fontSize/11) + spacingBefore + spacingAfter + (lines × fontSize × lineSpacing × 0.5)
```

**Tables:**
```
weight = (totalCharacters × 1.5) + (rows × 20) + 100
```

**List Items:**
```
weight = (characters × 1.2) + 30
```

**Images:**
```
weight = 800 (moderate default)
```

## Technical Details

### Key Functions

| Function | Purpose |
|----------|---------|
| `onOpen()` | Creates custom menu in Google Docs |
| `generateTableOfContents()` | Main function - orchestrates TOC generation |
| `getImprovedPageNumbers()` | Calculates page numbers using PDF export + weighting |
| `calculateElementWeight()` | Determines vertical space taken by each element |
| `estimateTOCPages()` | Predicts how many pages the TOC will occupy |

### Page Number Accuracy

The accuracy depends on:
- ✅ Consistent formatting throughout document
- ✅ Proper use of Heading 1 and Heading 2 styles
- ✅ Standard page margins and sizes
- ⚠️ May be slightly off if document has unusual formatting
- ⚠️ Multi-column layouts not fully supported

### Limitations

- Only supports Heading 1 and Heading 2 (can be extended)
- Assumes standard Google Docs page layouts
- Image weights use estimates (actual size not available via API)
- Very complex tables might have minor inaccuracies

## Customization

### Add Heading 3 Support

Find this section:
```javascript
if (heading === DocumentApp.ParagraphHeading.HEADING1) {
  // ...
} else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
  // ...
}
```

Add:
```javascript
else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
  headings.push({
    text: text,
    level: 3,
    index: i
  });
}
```

### Adjust TOC Formatting

Modify these sections in `generateTableOfContents()`:
- Line 165-187: Font sizes, spacing, bold/normal
- Line 151-154: Dot leader lengths and formatting

### Change Font Family

Line 161:
```javascript
newPara.setFontFamily('Roboto'); // Change to your preferred font
```

## Troubleshooting

### "Could not find Table of Contents heading"
- Ensure you have a paragraph containing "table of contents" (case-insensitive)

### "No headings found"
- Use Format → Paragraph styles → Heading 1 or Heading 2
- Don't just make text bold - use actual heading styles

### Page numbers are off by 1-2 pages
- Try running the generator twice (first run establishes TOC size)
- Very large documents may need slight manual adjustments

### Script times out
- Document may be too large (>100 pages with lots of tables)
- Try breaking into separate documents

## The Debug Output

Check **View → Logs** in Apps Script editor to see:
```
Total pages from PDF: 45
Total content weight: 125847.5
Page breaks detected: 3
Heading "Introduction..." at index 15 -> page 2 (weight: 1250.0/125847.5, proportion: 1.0%)
Heading "Methodology..." at index 89 -> page 15 (weight: 42500.0/125847.5, proportion: 33.8%)
```

## Credits

Built with precision and attention to detail to solve the challenge of accurate automated TOC generation in Google Docs.

**Key Insight**: The breakthrough was calculating page numbers BEFORE adding each heading's own weight, ensuring headings appear at their actual document position.

## License

Free to use and modify for your documents!

---

**Version**: 2.0 (Improved Accuracy)  
**Last Updated**: October 2025  
**Platform**: Google Apps Script for Google Docs

