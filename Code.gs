function onOpen() {
  DocumentApp.getUi()
    .createMenu('TOC Generator')
    .addItem('Update Table of Contents', 'generateTableOfContents')
    .addToUi();
}

function generateTableOfContents() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const ui = DocumentApp.getUi();
  
  // Find where to insert TOC (look for "Table of Contents" heading)
  let tocIndex = -1;
  
  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const text = para.getText().trim();
      
      if (text.toLowerCase().includes('table of contents')) {
        tocIndex = i;
        break;
      }
    }
  }
  
  if (tocIndex === -1) {
    ui.alert('Error', 'Could not find "Table of Contents" heading in your document.', ui.ButtonSet.OK);
    return;
  }
  
  // Delete existing TOC entries
  let deleteCount = 0;
  let safetyCounter = 0;
  
  while (tocIndex + 1 < body.getNumChildren() && safetyCounter < 500) {
    const nextElement = body.getChild(tocIndex + 1);
    safetyCounter++;
    
    if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = nextElement.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();
      
      // Stop if we hit the first main heading
      if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        break;
      }
      
      // Remove TOC-like entries
      if (text.includes('.....') || text === '' || /^\s*$/.test(text)) {
        body.removeChild(nextElement);
        deleteCount++;
        continue;
      }
      
      if (text.startsWith('\t') || /^\d+\./.test(text) || text.length < 5) {
        body.removeChild(nextElement);
        deleteCount++;
        continue;
      }
      
      break;
    } else {
      body.removeChild(nextElement);
      deleteCount++;
    }
  }
  
  console.log('Deleted ' + deleteCount + ' old TOC elements');
  
  // Collect all headings
  const headings = [];
  const currentNumChildren = body.getNumChildren();
  
  for (let i = 0; i < currentNumChildren; i++) {
    const element = body.getChild(i);
    
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();
      
      if (!text || i <= tocIndex) continue;
      
      if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        headings.push({
          text: text,
          level: 1,
          index: i
        });
      } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        headings.push({
          text: text,
          level: 2,
          index: i
        });
      }
    }
  }
  
  if (headings.length === 0) {
    ui.alert('No headings found', 'Make sure your sections use Heading 1 and Heading 2 styles.', ui.ButtonSet.OK);
    return;
  }
  
  // Get improved page numbers using weighted PDF estimation
  const { pageNumbers, totalContentPages } = getImprovedPageNumbers(doc, body, tocIndex);
  console.log('Total content pages from PDF: ' + totalContentPages);
  
  // Estimate TOC pages with better accuracy
  const tocPages = estimateTOCPages(headings, body, tocIndex);
  console.log('Estimated TOC will add ' + tocPages + ' pages');
  
  // Insert new TOC entries
  let insertIndex = tocIndex + 1;
  let previousLevel = 0;
  let isFirstMainSection = true;
  
  for (let i = 0; i < headings.length; i++) {
    const h = headings[i];
    
    // Get adjusted page number (content page + TOC shift)
    let basePage = pageNumbers[h.index] || 1;
    let pageNum = Math.max(1, basePage + tocPages);
    const pageNumStr = pageNum.toString();
    
    // Create the TOC entry text with dots
    let tocEntry = '';
    
    if (h.level === 1) {
      const dotsNeeded = Math.max(5, 65 - h.text.length - pageNumStr.length);
      tocEntry = h.text + '.'.repeat(dotsNeeded) + pageNumStr;
    } else {
      const dotsNeeded = Math.max(5, 60 - h.text.length - pageNumStr.length);
      tocEntry = '\t' + h.text + ' ' + '.'.repeat(dotsNeeded) + pageNumStr;
    }
    
    // Insert the paragraph
    const newPara = body.insertParagraph(insertIndex, tocEntry);
    
    // Style the paragraph
    newPara.setFontFamily('Roboto');
    newPara.setLineSpacing(1);
    
    if (h.level === 1) {
      newPara.setFontSize(12);
      newPara.editAsText().setBold(true);
      
      if (isFirstMainSection) {
        newPara.setSpacingBefore(0);
        isFirstMainSection = false;
      } else {
        newPara.setSpacingBefore(12);
      }
      
      newPara.setSpacingAfter(0);
    } else {
      newPara.setFontSize(11);
      newPara.editAsText().setBold(false);
      
      if (previousLevel === 1) {
        newPara.setSpacingBefore(6);
      } else {
        newPara.setSpacingBefore(0);
      }
      
      newPara.setSpacingAfter(6);
    }
    
    previousLevel = h.level;
    insertIndex++;
  }
  
  ui.alert('Success!', 'Table of Contents updated with ' + headings.length + ' entries (total pages: ' + totalContentPages + ', TOC pages: ' + tocPages + ').', ui.ButtonSet.OK);
}

/**
 * Estimate TOC pages more accurately
 */
function estimateTOCPages(headings, body, tocIndex) {
  // Get page dimensions and margins from document
  const pageHeight = body.getPageHeight();
  const pageWidth = body.getPageWidth();
  const marginTop = body.getMarginTop();
  const marginBottom = body.getMarginBottom();
  
  // Usable page height in points
  const usableHeight = pageHeight - marginTop - marginBottom;
  
  // Approximate heights for TOC entries
  const level1Height = 12 * 1.0 + 12; // font size * line spacing + spacing before
  const level2Height = 11 * 1.0 + 6;  // font size * line spacing + spacing after
  
  let totalHeight = 0;
  let isFirst = true;
  
  for (let i = 0; i < headings.length; i++) {
    if (headings[i].level === 1) {
      totalHeight += isFirst ? (12 * 1.0) : level1Height;
      isFirst = false;
    } else {
      totalHeight += level2Height;
    }
  }
  
  const estimatedPages = Math.ceil(totalHeight / usableHeight);
  return Math.max(1, estimatedPages);
}

/**
 * Calculate page numbers with improved accuracy considering formatting
 */
function getImprovedPageNumbers(doc, body, tocIndex) {
  // Get total pages from PDF export
  const blob = doc.getAs('application/pdf');
  const pdfData = blob.getDataAsString();
  
  let totalPages = 0;
  
  // Try multiple PDF page count methods
  const countMatch = pdfData.match(/\/N\s+(\d+)/);
  if (countMatch) {
    totalPages = parseInt(countMatch[1], 10);
  }
  
  if (totalPages === 0) {
    const re = /\/Count\s+(\d+)/g;
    let match;
    while ((match = re.exec(pdfData)) !== null) {
      let value = parseInt(match[1], 10);
      if (value > totalPages) totalPages = value;
    }
  }
  
  // Additional fallback
  if (totalPages === 0) {
    totalPages = (pdfData.match(/\/Page\b/g) || []).length;
  }
  
  console.log('Total pages from PDF: ' + totalPages);
  
  if (totalPages === 0) {
    console.error('Failed to extract page count from PDF');
    totalPages = 10; // Fallback estimate
  }
  
  // Calculate weighted page distribution
  const pageNumbers = {};
  let cumulativeWeight = 0;
  let totalWeight = 0;
  const numChildren = body.getNumChildren();
  const pageBreaks = [];
  
  // First pass: calculate total weight and detect page breaks
  for (let i = 0; i < numChildren; i++) {
    if (i <= tocIndex) continue; // Skip TOC section
    
    const element = body.getChild(i);
    const weight = calculateElementWeight(element);
    totalWeight += weight;
    
    // Check for page breaks
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const attributes = para.getAttributes();
      if (attributes[DocumentApp.Attribute.PAGE_BREAK_BEFORE]) {
        pageBreaks.push(i);
      }
    }
  }
  
  console.log('Total content weight: ' + totalWeight);
  console.log('Page breaks detected: ' + pageBreaks.length);
  
  // Second pass: assign page numbers to headings
  let currentPage = 1;
  let pageBreakIndex = 0;
  cumulativeWeight = 0; // Reset for second pass
  
  for (let i = 0; i < numChildren; i++) {
    if (i <= tocIndex) continue;
    
    const element = body.getChild(i);
    
    // Check if we passed a page break BEFORE processing this element
    if (pageBreakIndex < pageBreaks.length && i >= pageBreaks[pageBreakIndex]) {
      currentPage++;
      pageBreakIndex++;
    }
    
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const heading = para.getHeading();
      
      if (heading === DocumentApp.ParagraphHeading.HEADING1 || 
          heading === DocumentApp.ParagraphHeading.HEADING2) {
        
        // Calculate page BEFORE adding this heading's weight
        // The heading appears at this position, not after its own content
        const proportion = cumulativeWeight / totalWeight;
        const proportionalPage = Math.max(1, Math.ceil(proportion * totalPages));
        
        // Use the greater of proportional or explicit page break count
        pageNumbers[i] = Math.max(currentPage, proportionalPage);
        
        console.log('Heading "' + para.getText().substring(0, 30) + '..." at index ' + i + 
                    ' -> page ' + pageNumbers[i] + 
                    ' (weight: ' + cumulativeWeight.toFixed(0) + '/' + totalWeight.toFixed(0) + 
                    ', proportion: ' + (proportion * 100).toFixed(1) + '%)');
      }
    }
    
    // Add weight AFTER we've assigned page numbers for this element
    const weight = calculateElementWeight(element);
    cumulativeWeight += weight;
  }
  
  return { pageNumbers, totalContentPages: totalPages };
}

/**
 * Calculate weight of an element considering formatting
 */
function calculateElementWeight(element) {
  const type = element.getType();
  
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    const para = element.asParagraph();
    const text = para.getText();
    const numChars = text.length;
    
    if (numChars === 0) return 20; // Empty paragraph
    
    // Get font size (affects space taken)
    let fontSize = 11; // Default
    try {
      const textElement = para.editAsText();
      if (textElement.getFontSize(0)) {
        fontSize = textElement.getFontSize(0);
      }
    } catch (e) {
      // Use default if can't read font size
    }
    
    // Get spacing
    let spacingBefore = para.getSpacingBefore() || 0;
    let spacingAfter = para.getSpacingAfter() || 0;
    let lineSpacing = para.getLineSpacing() || 1.15;
    
    // Calculate lines needed (approximate)
    const avgCharsPerLine = Math.floor(500 / fontSize); // Rough estimate
    const lines = Math.max(1, Math.ceil(numChars / avgCharsPerLine));
    
    // Weight = character count * font size factor + spacing + line breaks
    const baseWeight = numChars * (fontSize / 11.0);
    const spacingWeight = spacingBefore + spacingAfter;
    const lineWeight = lines * fontSize * lineSpacing * 0.5;
    
    return baseWeight + spacingWeight + lineWeight;
    
  } else if (type === DocumentApp.ElementType.TABLE) {
    const table = element.asTable();
    const numRows = table.getNumRows();
    let totalCells = 0;
    let totalChars = 0;
    
    for (let r = 0; r < numRows; r++) {
      const row = table.getRow(r);
      const numCells = row.getNumCells();
      totalCells += numCells;
      
      for (let c = 0; c < numCells; c++) {
        const cell = row.getCell(c);
        totalChars += cell.getText().length;
      }
    }
    
    // Tables take up more space than regular text
    const tableWeight = totalChars * 1.5 + numRows * 20 + 100;
    return tableWeight;
    
  } else if (type === DocumentApp.ElementType.LIST_ITEM) {
    const listItem = element.asListItem();
    const text = listItem.getText();
    // List items have indentation and bullet points
    return text.length * 1.2 + 30;
    
  } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    // Images can vary greatly, use a moderate default
    return 800;
  }
  
  // Other element types
  return 50;
}
