# Missing Features in Python Implementation

## Critical Missing Methods

### 1. `renderCommentContent` (Used in `renderCommentReference`)
**Location**: TypeScript line 1026-1031
**Purpose**: Renders the full comment content including author, date, and comment body
**Status**: ❌ Missing
**Impact**: Comment references won't show full comment details

### 2. `later` (Deferred Task Execution)
**Location**: TypeScript line 1553-1555
**Purpose**: Defers function execution to post-render phase
**Status**: ❌ Missing
**Impact**: Comment range start/end markers won't work correctly (uses Range API)

### 3. `renderRunProperties` (Wrapper Method)
**Location**: TypeScript line 924-926
**Purpose**: Wrapper that calls `renderCommonProperties` for run properties
**Status**: ❌ Missing (but may not be used directly)

## Incomplete Implementations

### 4. `renderCommentReference` - Incomplete
**Current**: Only returns a simple span with emoji
**Expected**: Should return DocumentFragment-like structure with:
- Comment marker (💬)
- Comment popover container with:
  - Author name
  - Date
  - Comment content (rendered via `renderCommentContent`)
**Impact**: Comments won't display full information on hover/click

### 5. `renderCommentRangeStart` - Missing Range API Integration
**Current**: Only creates comment node
**Expected**: Should:
- Create Range object
- Add to commentHighlight (CSS Highlight API)
- Use `later()` to set range start after DOM is ready
- Store in commentMap
**Impact**: Comment highlighting won't work

### 6. `renderCommentRangeEnd` - Missing Range API Integration
**Current**: Only creates comment node
**Expected**: Should:
- Retrieve Range from commentMap
- Use `later()` to set range end after DOM is ready
**Impact**: Comment highlighting won't work

## Browser-Specific Features (May Not Apply to Python)

### 7. CSS Highlight API Integration
**Location**: TypeScript lines 87-89, 142-144
**Purpose**: Uses browser's CSS Highlight API for comment ranges
**Status**: ⚠️ Browser-specific, may not be applicable
**Note**: Python/BeautifulSoup doesn't have browser APIs

### 8. VML SVG Dynamic Sizing
**Location**: TypeScript lines 1272-1277
**Purpose**: Uses `requestAnimationFrame` and SVG `getBBox()` to calculate dimensions
**Status**: ⚠️ Partially missing - Python version doesn't calculate SVG dimensions
**Impact**: SVG containers may not have correct width/height attributes

## Convenience Methods (Low Priority)

### 9. `createSvgElement`
**Location**: TypeScript line 1541-1543
**Purpose**: Convenience wrapper for `createElementNS` with SVG namespace
**Status**: ⚠️ Not needed - Python can use `create_element_ns` directly
**Impact**: None - just a convenience method

## Summary

**✅ IMPLEMENTED**:
1. ✅ `renderCommentContent` - Comment content rendering
2. ✅ `later` - Deferred task execution  
3. ✅ Complete `renderCommentReference` implementation
4. ✅ `renderRunProperties` (wrapper method)
5. ✅ `renderCommentRangeStart` / `renderCommentRangeEnd` - Basic implementation (without browser Range API)

**Browser-Specific** (may not apply - skipped):
- CSS Highlight API integration
- Range API for comment markers (basic placeholder implemented)
- `requestAnimationFrame` for VML sizing

**Low Priority** (not needed):
- `createSvgElement` (convenience method - Python uses `create_element_ns` directly)

