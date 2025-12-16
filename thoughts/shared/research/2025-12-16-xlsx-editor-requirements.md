---
date: 2025-12-16T11:17:25+10:00
researcher: rocky
git_commit: no-git-history
branch: no-branch
repository: xlsx
topic: "XLSX File Editor with ExcelJS - Requirements and Architecture Research"
tags: [research, requirements, exceljs, xlsx, browser-app, inventory-management]
status: complete
last_updated: 2025-12-16
last_updated_by: rocky
last_updated_note: "Added GitHub Pages hosting decision"
---

# Research: XLSX File Editor with ExcelJS - Requirements and Architecture

**Date**: 2025-12-16T11:17:25+10:00
**Researcher**: rocky
**Git Commit**: no-git-history (new project)
**Branch**: no-branch
**Repository**: xlsx

## Research Question

Document requirements for a browser-based xlsx file editor using ExcelJS with inventory management features including SKU search, quantity editing, date tracking, low-stock highlighting, and sync capabilities.

## Summary

This is a **new project** with no existing codebase. The directory currently contains only Claude Code configuration files. This document captures the functional requirements for building a browser-based spreadsheet application for inventory management.

## Current Project State

### Directory Structure
```
/home/rocky/web/xlsx/
├── .claude/
│   ├── agents/          # Sub-agent configurations
│   ├── commands/        # Slash command templates
│   └── settings.json    # Claude Code settings
└── .gitignore           # Git ignore configuration
```

### Existing Files
- `.gitignore` - Excludes `.claude/settings.local.json`
- `.claude/` - Claude Code tooling configuration (not application code)

**No application code exists yet.**

## Detailed Requirements

### 1. Core Technology Stack

| Component | Technology |
|-----------|------------|
| Spreadsheet Library | ExcelJS (via CDN) |
| Runtime Environment | Browser-based (client-side only) |
| File Format | .xlsx (Excel format) |
| **Hosting** | **GitHub Pages (static site)** |
| UI | Plain HTML/CSS/JS (no build step) |

### 2. Feature Requirements

#### 2.1 Search Functionality
- **Input**: SKU code (product identifier)
- **Output**: Quantity of that product
- **Behavior**: Search through spreadsheet data to find matching SKU and return associated quantity

#### 2.2 Edit Functionality
- **Primary Action**: Edit quantity value for a product
- **Secondary Actions**:
  - Add current date to a designated cell (timestamp of edit)
  - Clear/clean the next cell (possibly for notes or status)
- **Trigger**: User-initiated edit operation

#### 2.3 Visual Highlighting
- **Purpose**: Mark low-quantity items for ordering
- **Style**: Yellow background color
- **Trigger**: Manual user action (not automatic threshold)
- **Use Case**: Visual indicator that item needs to be reordered

#### 2.4 Update/Header Functionality
- **Location**: Top of first sheet/page
- **Content**: Name and date
- **Purpose**: Document who updated the file and when

#### 2.5 Sync Functionality
- **Option 1**: OneDrive cloud synchronization
- **Option 2**: USB stick / local device sync
- **Purpose**: Backup and cross-device access

#### 2.6 Testing
- **Requirement**: Comprehensive test suite
- **Coverage**: All major features (search, edit, highlight, update, sync)

## Architecture Considerations

### GitHub Pages Hosting

**Decision**: Host as a static site on GitHub Pages.

**Implications**:
- No server-side code (no Node.js, no backend)
- All processing happens in the browser (client-side)
- ExcelJS loaded via CDN (e.g., cdnjs, unpkg, or jsDelivr)
- No build step required - plain HTML/CSS/JS files
- Free hosting with HTTPS
- Easy deployment via `gh-pages` branch or `/docs` folder

**Project Structure for GitHub Pages**:
```
/home/rocky/web/xlsx/
├── index.html          # Main entry point
├── css/
│   └── style.css       # Styles
├── js/
│   └── app.js          # Application logic
├── tests/
│   └── *.test.js       # Test files
└── README.md           # Project documentation
```

**CDN for ExcelJS**:
```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
```

### Browser-Based ExcelJS Usage

ExcelJS can run in the browser with some considerations:
- File reading via `FileReader` API
- File writing via `Blob` and download trigger
- No direct filesystem access (browser security sandbox)
- All data stays client-side (privacy advantage)

### Data Flow

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│  Load File  │────>│   ExcelJS    │────>│  In-Memory  │
│  (Upload)   │     │   Parser     │     │  Workbook   │
└─────────────┘     └──────────────┘     └─────────────┘
                                               │
                    ┌──────────────────────────┘
                    ▼
              ┌───────────┐
              │  Search   │──> Find SKU, return quantity
              │  Edit     │──> Modify cells, add date
              │  Highlight│──> Apply yellow background
              │  Update   │──> Modify header row
              └───────────┘
                    │
                    ▼
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│  Save File  │<────│   ExcelJS    │<────│  Modified   │
│  (Download) │     │   Writer     │     │  Workbook   │
└─────────────┘     └──────────────┘     └─────────────┘
```

### Sync Architecture Options

#### OneDrive Integration
- Microsoft Graph API for OneDrive access
- OAuth2 authentication required
- **Challenge for static site**: OAuth flow typically needs a backend for token exchange
- **Workaround options**:
  - Use MSAL.js (Microsoft Authentication Library) for client-side auth
  - Or simply use OneDrive's folder sync on desktop (user saves file to synced folder)
- Can read/write files directly to cloud if authenticated

#### USB/Local Device (Simpler for Static Site)
- File System Access API (modern browsers like Chrome)
- Or traditional download/upload workflow
- User manually manages file placement to USB
- **Recommended for GitHub Pages**: Download modified file, user saves to desired location

## Open Questions

1. **Spreadsheet Structure**: What is the expected column layout? (SKU column, Quantity column, Date column positions)
2. **Search Scope**: Search all sheets or specific sheet only?
3. **Multiple Results**: How to handle if SKU appears multiple times?
4. **Date Format**: Preferred date format for the date cell?
5. **Header Format**: Exact format for name/date in header?
6. **Sync Priority**: Which sync method is primary (OneDrive vs USB)?
7. **Authentication**: How to handle OneDrive authentication flow? (Note: complex for static site)
8. **Offline Support**: Should the app work offline? (Service Worker for GitHub Pages?)
9. ~~**UI Framework**: Plain HTML/JS or use a framework?~~ **ANSWERED**: Plain HTML/CSS/JS (no build step)
10. **Test Framework**: Browser-based testing (e.g., QUnit, Mocha in browser) since no Node.js?

## Implementation Phases (Suggested)

### Phase 1: Core Functionality
- File upload/download with ExcelJS
- Basic spreadsheet display
- SKU search functionality

### Phase 2: Editing Features
- Quantity editing
- Date cell population
- Cell clearing

### Phase 3: Visual Features
- Yellow highlight for low stock
- Header update functionality

### Phase 4: Sync & Polish
- OneDrive integration OR local file handling
- Test suite implementation

## Technical References

### ExcelJS Resources
- NPM: `npm install exceljs`
- Browser bundle available via CDN
- Documentation: https://github.com/exceljs/exceljs

### Key ExcelJS APIs
```javascript
// Reading
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(buffer);

// Accessing cells
const worksheet = workbook.getWorksheet(1);
const cell = worksheet.getCell('A1');

// Styling (yellow background)
cell.fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF00' }
};

// Writing
const buffer = await workbook.xlsx.writeBuffer();
```

## Related Research

No existing research documents in this project.

## Next Steps

1. Answer open questions to finalize requirements
2. Create implementation plan with `/create_plan`
3. Set up project structure (package.json, HTML entry point)
4. Begin Phase 1 implementation
