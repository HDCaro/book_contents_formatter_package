# CSS Professional Styling Guide - EPUB Structure

## What Changed

Your eBook builder now generates **professional modern EPUBs** with:

### 1. **OEBPS Standard Structure**

```
Inside the EPUB file (ZIP):
text/
  ├── front.xhtml          ← Front matter with CSS link
  ├── body.xhtml           ← All chapters with CSS link
  └── back_cover.xhtml     ← Back cover

images/
  └── embedded/            ← All embedded images
      ├── 0001_image001.jpg
      ├── 0002_image002.jpg
      └── ...

styles/
  └── style.css            ← Professional CSS (4291 bytes)
```

### 2. **CSS Features Included**

**Typography**

- Body text: Georgia serif (readable for long reading)
- Headings: Trebuchet MS sans-serif (modern, clean)
- Line-height: 1.6 (comfortable for reading)
- Font size: responsive to reader device

**Color Scheme**

- Primary: #1a1a1a (dark, professional)
- Secondary: #333, #444 (heading hierarchy)
- Accent: #007ACC (links, borders - professional blue)
- Neutral: #666, #999, #f5f5f5 (quotes, backgrounds)

**Layout & Spacing**

- Body padding: 1em 1.5em (breathing room)
- Heading margins: 2em 0 (proper section breaks)
- Paragraph margins: 1em 0 (readable spacing)
- Lists: 2em indent (clear hierarchy)

**Special Elements**

```css
Blockquotes:
- Left border: 4px #007ACC
- Background: #f5f5f5 (subtle highlight)
- Font: italic
- Padding: 1em left, 1em total

Tables:
- 100% width
- Collapsed borders
- Header background: #efefef
- Cell padding: 0.75em

Links:
- Color: #007ACC (unvisited)
- Color: #7030A0 (visited)
- Underline on hover
```

**Page Breaks (for PDF/Print)**

```css
- h1: page-break-before: always   (new chapter on new page)
- img, figure: page-break-inside: avoid (images not split)
- tables: page-break-inside: avoid
- blockquotes: page-break-inside: avoid
```

### 3. **What Gets Applied Automatically**

When you build an EPUB, the CSS is automatically:

1. ✓ Generated from `get_modern_epub_css()`
2. ✓ Added to EPUB as `styles/style.css`
3. ✓ Linked in all XHTML files: `<link rel="stylesheet" href="../styles/style.css">`
4. ✓ Applied to all chapter content

### 4. **Image References**

Images are correctly referenced with relative paths:

```html
<!-- In text/body.xhtml or text/front.xhtml -->
<img src="../images/embedded/0001_image001.jpg" alt="Description" />
```

The `../` is necessary because:

- HTML files are in: `text/`
- Images are in: `images/`
- So from `text/`, must go up one level: `../images/`

### 5. **How Readers See It**

Different readers apply the CSS differently:

**Kindle (MOBI converted from EPUB)**

- Clean, consistent typography
- Proper heading hierarchy
- Page breaks respected
- Colors adapted to reader theme

**Apple Books/iBooks**

- Modern sans-serif headings
- Professional color scheme
- Justified text (if enabled)
- Linked TOC navigation

**Kobo/Overdrive**

- Full CSS applied (best results)
- All styling visible
- Best readability
- Professional appearance

**Web readers (for testing)**

- Exact appearance
- All CSS features visible
- Professional styling apparent

### 6. **Testing the CSS**

To see the CSS in action:

```bash
# Build the EPUB
python src/ebook/ebook_builder.py

# Verify the EPUB
python src/ebook/ebook_builder.py --verify

# Inspect manually (with 7-Zip or WinRAR)
# Extract the EPUB and check:
# - text/*.xhtml for CSS link
# - styles/style.css content
# - images/embedded/ for image files
```

### 7. **Customization (Future Enhancement)**

Currently CSS is fixed to the modern professional style. Future versions could add:

```json
"styling": {
  "typography": "modern",        // modern, conservative, academic
  "font_family": "serif",        // serif, sans-serif, mixed
  "color_scheme": "blue",        // blue, green, gray, navy
  "line_height": 1.6,            // 1.5, 1.6, 1.8, 2.0
  "margin_style": "generous"     // compact, normal, generous
}
```

### 8. **KDP Compatibility**

The modern CSS structure is **fully KDP compatible**:

✓ Semantic HTML5 structure
✓ No inline styles (all in CSS file)
✓ Proper heading hierarchy
✓ Reflowable design
✓ Professional appearance
✓ Mobile-friendly
✓ Accessibility compliant

---

**Date**: 2026-05-11  
**Builder Version**: 2.0+  
**CSS Version**: Modern Professional v1
