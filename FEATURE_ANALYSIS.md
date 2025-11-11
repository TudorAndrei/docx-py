# TypeScript docxjs Feature-by-Feature Analysis

## Overview
This document provides a comprehensive feature-by-feature analysis of the TypeScript docxjs codebase to identify all capabilities and ensure the Python port matches functionality.

---

## 1. CORE RENDERING FEATURES

### 1.1 Document Structure
- **Document Element**: Root container (`DomType.Document`)
- **Sections**: Document sections with properties (margins, page size, orientation)
- **Page Breaking**: 
  - Manual page breaks (`<w:br w:type="page"/>`)
  - Application page breaks (`<w:lastRenderedPageBreak/>`)
  - Section-based page breaks (orientation/size changes)
- **Page Layout**: Width, height, margins, orientation (portrait/landscape)
- **Background**: Document background color support

### 1.2 Paragraphs (`WmlParagraph`)
- **Basic Rendering**: Paragraphs as `<p>` elements
- **Paragraph Properties**:
  - Text alignment (left, center, right, justify, etc.)
  - Line spacing (single, double, multiple, exact, at-least)
  - Indentation (left, right, first-line, hanging)
  - Spacing (before, after)
  - Keep lines together (`keepLines`)
  - Keep with next (`keepNext`)
  - Page break before (`pageBreakBefore`)
  - Outline level (`outlineLevel`)
  - Borders (top, bottom, left, right)
  - Shading/background color
- **Tab Stops**: 
  - Tab positions and styles (left, right, center, decimal, bar)
  - Tab leaders (none, dot, hyphen, underscore, etc.)
  - Experimental tab stop calculation (requires `experimental: true`)
- **Section Properties**: Paragraphs can contain section breaks
- **Numbering**: Paragraph-level numbering/list support

### 1.3 Text Runs (`WmlRun`)
- **Basic Rendering**: Runs as `<span>` elements
- **Run Properties**:
  - Font family, size, color
  - Bold, italic, underline, strikethrough
  - Text effects (shadow, outline, emboss, engrave)
  - Character spacing
  - Vertical alignment (superscript, subscript)
  - Highlight color
  - Shading
  - Borders
- **Field Runs**: Special handling for field content
- **Vertical Alignment**: Wrapper elements for superscript/subscript

### 1.4 Text Content
- **Plain Text** (`WmlText`): Regular text nodes
- **Deleted Text** (`WmlDeletedText`): Tracked deletions (requires `renderChanges: true`)
- **No-Break Hyphen**: Special hyphen character
- **Symbols** (`WmlSymbol`): Unicode symbols with font specification
- **Tabs**: Tab characters (rendered as non-breaking spaces or styled spans)
- **Line Breaks** (`WmlBreak`):
  - Text wrapping breaks (`<br>`)
  - Page breaks (creates new page)
  - Last rendered page breaks (optional, controlled by `ignoreLastRenderedPageBreak`)

---

## 2. STYLING SYSTEM

### 2.1 Document Styles (`IDomStyle`)
- **Style Types**: Paragraph styles, character styles, table styles, list styles
- **Style Hierarchy**: 
  - Based-on relationships (`basedOn`)
  - Linked styles (`linked`)
  - Next styles (`next`)
- **Style Properties**: Merges paragraph and run properties
- **Default Styles**: Document default styles for paragraphs and runs
- **Style Processing**: 
  - Style name processing with class name prefix
  - CSS class generation (`className_styleName`)
  - Style inheritance resolution

### 2.2 Theme Support
- **Color Scheme**: 
  - Theme colors (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
  - CSS custom properties (`--docx-*-color`)
- **Font Scheme**:
  - Major font (headings)
  - Minor font (body)
  - CSS custom properties (`--docx-majorHAnsi-font`, `--docx-minorHAnsi-font`)
- **Theme Variables**: Injected as CSS variables in `.docx` selector

### 2.3 Font Table
- **Embedded Fonts**: Font file loading and `@font-face` generation
- **Font Variants**: Regular, bold, italic, bold-italic
- **Font Loading**: Async font loading with Promise-based tasks
- **Font Family**: Font name wrapping for CSS compatibility

### 2.4 Common Properties
- **Margins**: Top, bottom, left, right
- **Borders**: All sides with style, width, color
- **Shading**: Background color/pattern
- **Spacing**: Before, after spacing
- **Indentation**: Left, right, first-line, hanging

---

## 3. TABLES

### 3.1 Table Structure (`WmlTable`)
- **Basic Rendering**: Tables as `<table>` elements
- **Table Columns**: Column width specification (`<colgroup>`)
- **Table Rows** (`WmlTableRow`):
  - Header rows (`isHeader`)
  - Grid before/after (hidden columns)
- **Table Cells** (`WmlTableCell`):
  - Cell spanning (`colspan`, `rowspan`)
  - Vertical merging (`verticalMerge: 'restart' | 'continue'`)
  - Cell borders and padding
  - Cell background/shading
  - Vertical alignment

### 3.2 Table Styling
- **Cell Styles**: Default cell style application
- **Table Borders**: Table-level border settings
- **Column Banding**: Alternating column styles (`colBandSize`)
- **Row Banding**: Alternating row styles (`rowBandSize`)

---

## 4. NUMBERING AND LISTS

### 4.1 Numbering System
- **Numbering Definitions**: Abstract and concrete numbering definitions
- **Numbering Levels**: Multi-level numbering support
- **Numbering Formats**:
  - Decimal, lower/upper roman, lower/upper letter
  - Bullet (custom images)
  - Custom formats (Chinese, Japanese, Korean, etc.)
- **Numbering Styles**: Linked to paragraph styles

### 4.2 List Rendering
- **CSS Counters**: Uses CSS `counter-increment` and `counter-reset`
- **Bullet Images**: Custom bullet images loaded asynchronously
- **Level Text**: Custom numbering text with placeholders (`%1`, `%2`, etc.)
- **List Styling**: List-style-type, list-style-position

### 4.3 Numbering Properties
- **Start Value**: Custom starting numbers
- **Restart**: Restart numbering at specific levels
- **Suffix**: Text after number (tab, space, none)
- **Justification**: Number alignment

---

## 5. HEADERS AND FOOTERS

### 5.1 Header/Footer Types
- **Default**: Standard header/footer
- **First Page**: Special first page header/footer
- **Even/Odd Pages**: Different headers/footers for even/odd pages

### 5.2 Header/Footer Rendering
- **Positioning**: Margin-based positioning relative to page margins
- **Content**: Full document content support (paragraphs, tables, images, etc.)
- **Section-Based**: Each section can have different headers/footers
- **Page Numbering**: Automatic page number insertion (via fields)

---

## 6. FOOTNOTES AND ENDNOTES

### 6.1 Note References
- **Footnote References**: Inline superscript references
- **Endnote References**: Similar to footnotes but rendered at end
- **Note Numbering**: Automatic sequential numbering

### 6.2 Note Content
- **Note Rendering**: Notes rendered as `<ol>` lists
- **Note Placement**: 
  - Footnotes at bottom of each page
  - Endnotes at end of document (last page only)
- **Note Content**: Full document content support

---

## 7. COMMENTS (Experimental)

### 7.1 Comment System
- **Comment References**: Inline comment markers
- **Comment Ranges**: Start/end markers for commented text
- **Comment Content**: Author, date, initials, text content
- **Comment Highlighting**: Uses CSS Custom Highlight API (if available)

### 7.2 Comment Rendering
- **Comment Popovers**: Hover/click to view comments
- **Comment Styling**: Special CSS classes for comment elements
- **Comment Map**: Tracks comment ranges for highlighting

---

## 8. HYPERLINKS AND BOOKMARKS

### 8.1 Hyperlinks (`WmlHyperlink`)
- **URL Links**: External links (`href`)
- **Anchor Links**: Internal document links (`anchor`)
- **Link Styling**: Inherited text styling

### 8.2 Bookmarks (`WmlBookmarkStart`, `WmlBookmarkEnd`)
- **Bookmark Anchors**: Creates `<span id="bookmarkName">` elements
- **Internal Navigation**: Enables anchor-based navigation

---

## 9. IMAGES AND DRAWINGS

### 9.1 Images (`IDomImage`)
- **Image Loading**: Async image loading from document parts
- **Image Properties**:
  - Source rect (cropping)
  - Rotation
  - Transform (scale, rotate)
  - Clip path for cropping
- **Image Rendering**: `<img>` elements with styling

### 9.2 Drawings (`Drawing`)
- **Drawing Container**: `<div>` with inline-block positioning
- **Drawing Content**: Can contain images, shapes, etc.

### 9.3 VML Elements (`VmlElement`)
- **VML Rendering**: Vector graphics as SVG
- **VML Pictures**: Embedded images in VML
- **VML Shapes**: Complex shapes and paths

---

## 10. EQUATIONS (MathML)

### 10.1 MathML Support
- **Math Elements**: Full MathML namespace support
- **Equation Types**:
  - Fractions (`mfrac`)
  - Radicals (`mroot`, `msqrt`)
  - Superscripts/Subscripts (`msup`, `msub`)
  - Matrices (`mtable`, `mtr`, `mtd`)
  - Functions (`mfunc`)
  - Delimiters (`mdel`)
  - And many more...

### 10.2 Math Rendering
- **MathML Namespace**: Proper MathML namespace handling
- **Equation Arrays**: Multi-line equations
- **Group Characters**: Over/under braces, brackets

---

## 11. FIELDS

### 11.1 Field Types
- **Simple Fields** (`WmlFieldSimple`): Inline field instructions
- **Complex Fields** (`WmlFieldChar`): Field begin/end/separate markers
- **Field Instructions**: Text instructions for field evaluation

### 11.2 Field Rendering
- **Field Handling**: Currently not fully rendered (field evaluation not implemented)
- **Field Runs**: Special run type for field content

---

## 12. TRACKED CHANGES (Experimental)

### 12.1 Change Tracking
- **Insertions** (`Inserted`): Rendered as `<ins>` elements
- **Deletions** (`Deleted`): Rendered as `<del>` elements
- **Deleted Text**: Strikethrough text (if `renderChanges: false`)

### 12.2 Change Rendering
- **Optional Rendering**: Controlled by `renderChanges` option
- **Change Styling**: CSS classes for inserted/deleted content

---

## 13. ALT CHUNKS

### 13.1 Alt Chunk Support
- **HTML Parts**: Embedded HTML content in documents
- **Alt Chunk Loading**: Async loading of HTML parts
- **Alt Chunk Rendering**: `<iframe>` with `srcdoc` attribute

---

## 14. SMART TAGS

### 14.1 Smart Tag Support
- **Smart Tag Elements**: XML smart tag elements
- **Smart Tag Rendering**: Rendered as `<span>` elements

---

## 15. DOCUMENT PROPERTIES

### 15.1 Core Properties
- **Metadata**: Title, author, subject, keywords, etc.
- **Creation/Modification Dates**: Document timestamps

### 15.2 Extended Properties
- **Application Info**: Application name, version
- **Template Info**: Template information
- **Statistics**: Word count, page count, etc.

### 15.3 Custom Properties
- **Custom Metadata**: User-defined properties

---

## 16. PAGE LAYOUT AND SECTIONS

### 16.1 Section Properties
- **Page Size**: Width and height
- **Page Margins**: Top, bottom, left, right, header, footer
- **Page Orientation**: Portrait or landscape
- **Columns**: Multi-column layout support
- **Header/Footer References**: Section-specific headers/footers

### 16.2 Section Rendering
- **Section Elements**: `<section>` elements for pages
- **Section Content**: `<article>` elements for column content
- **Page Breaking**: Section-based page breaks

---

## 17. EXPERIMENTAL FEATURES

### 17.1 Tab Stop Calculation
- **Experimental Tab Stops**: Dynamic tab stop width calculation
- **Pixel-to-Point Conversion**: Browser-based measurement
- **Tab Stop Styling**: Uses `word-spacing` CSS property

### 17.2 Comment Highlighting
- **CSS Custom Highlight**: Uses Highlight API for comment ranges
- **Comment Map**: Tracks comment ranges in DOM

---

## 18. RENDERING OPTIONS

### 18.1 Display Options
- `inWrapper`: Wrap document in container div
- `hideWrapperOnPrint`: Hide wrapper on print
- `ignoreWidth`: Ignore page width constraints
- `ignoreHeight`: Ignore page height constraints
- `breakPages`: Enable page breaking
- `ignoreLastRenderedPageBreak`: Ignore application page breaks

### 18.2 Content Options
- `renderHeaders`: Enable header rendering
- `renderFooters`: Enable footer rendering
- `renderFootnotes`: Enable footnote rendering
- `renderEndnotes`: Enable endnote rendering
- `renderComments`: Enable comment rendering (experimental)
- `renderChanges`: Enable tracked changes rendering (experimental)
- `renderAltChunks`: Enable alt chunk rendering

### 18.3 Styling Options
- `ignoreFonts`: Disable font loading
- `className`: CSS class name prefix
- `useBase64URL`: Use base64 for images/fonts instead of blob URLs

### 18.4 Technical Options
- `trimXmlDeclaration`: Remove XML declarations before parsing
- `experimental`: Enable experimental features
- `debug`: Enable debug logging

---

## 19. ASYNC OPERATIONS

### 19.1 Async Resource Loading
- **Font Loading**: Async font file loading
- **Image Loading**: Async image file loading
- **Numbering Images**: Async bullet image loading
- **Alt Chunks**: Async HTML part loading

### 19.2 Task Management
- **Promise-based Tasks**: All async operations collected as Promises
- **Task Execution**: `Promise.allSettled()` for all tasks
- **Post-render Tasks**: Tasks executed after initial render

---

## 20. UTILITIES AND HELPERS

### 20.1 XML Parsing
- **XML Parser**: Custom XML parser with namespace support
- **Element/Attribute Access**: Local name-based access
- **Length Conversion**: DXA, EMU, point conversions

### 20.2 CSS Generation
- **Style String Generation**: CSS rule generation from objects
- **Class Name Escaping**: CSS-safe class name generation
- **Font Family Wrapping**: CSS font-family string generation

### 20.3 DOM Utilities
- **Element Creation**: Helper methods for creating DOM elements
- **Namespace Support**: SVG and MathML namespace handling
- **Style Application**: CSS style object to string conversion

---

## 21. MISSING/NOT IMPLEMENTED FEATURES

### 21.1 Fields
- **Field Evaluation**: Field instructions not evaluated
- **Table of Contents**: TOC fields not supported
- **Page Numbers**: Page number fields not dynamically calculated

### 21.2 Real-time Features
- **Real-time Page Breaking**: Not implemented (performance concerns)
- **Dynamic Sizing**: No re-calculation on content changes

### 21.3 Advanced Features
- **Content Controls**: SDT (Structured Document Tags) partially supported
- **Form Fields**: Form field rendering not implemented
- **Macros**: VBA macros not supported

---

## SUMMARY

The TypeScript docxjs library is a comprehensive DOCX to HTML renderer with:
- **47 source files** covering all major DOCX features
- **64 DOM element types** for different document elements
- **25+ rendering options** for customization
- **Full styling support** including themes, fonts, and styles
- **Advanced features** like equations, comments, tracked changes
- **Experimental features** for tab stops and comment highlighting

The Python port should match all these features to ensure compatibility and feature parity.

