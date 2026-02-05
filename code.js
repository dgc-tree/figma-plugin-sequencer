// ============================================================
// SEQUENCER - Figma Plugin (Compliance-First Design)
// ============================================================
// A sequential numbering system for Figma with two modes:
//
// COMPLIANCE MODE - For invoices, quotes, binding documents
//   - Numbers are sequential and unique
//   - Cannot change prefix after first use
//   - Cannot reset to lower than highest used
//   - Cannot re-stamp already stamped layers
//
// DESIGN MODE - For prototypes and wireframes
//   - Fully flexible, can reset anytime
//   - Can change prefix, re-stamp, etc.
// ============================================================

// ----------------------------------------------------------
// STORAGE KEYS
// ----------------------------------------------------------
const SEQUENCES_KEY = 'sequences';
const SELECTED_KEY = 'selectedSequence';
const MIGRATION_KEY = 'migrationVersion';
const CURRENT_VERSION = 2;

// ----------------------------------------------------------
// STORAGE HELPERS
// ----------------------------------------------------------
function getSequences() {
  const data = figma.root.getPluginData(SEQUENCES_KEY);
  if (!data) return [];
  try {
    return JSON.parse(data);
  } catch (e) {
    console.error('Failed to parse sequences:', e);
    return [];
  }
}

function saveSequences(sequences) {
  figma.root.setPluginData(SEQUENCES_KEY, JSON.stringify(sequences));
}

function getSelectedId() {
  return figma.root.getPluginData(SELECTED_KEY) || null;
}

function setSelectedId(id) {
  figma.root.setPluginData(SELECTED_KEY, id || '');
}

function generateId() {
  return 'seq_' + Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

// ----------------------------------------------------------
// MIGRATION
// ----------------------------------------------------------
function migrateIfNeeded() {
  const version = figma.root.getPluginData(MIGRATION_KEY);
  const versionNum = version ? parseInt(version) : 0;

  if (versionNum < CURRENT_VERSION) {
    const sequences = getSequences();

    // Migrate v1 sequences to v2 (add mode, highestUsed, locked)
    sequences.forEach(seq => {
      if (!seq.mode) {
        seq.mode = 'compliance'; // Default to compliance
        seq.highestUsed = seq.value || seq.nextValue || '0000';
        seq.nextValue = seq.value || '0001';
        seq.locked = false;
        seq.createdAt = Date.now();
        delete seq.value; // Rename to nextValue
      }
    });

    if (sequences.length > 0) {
      saveSequences(sequences);
    }

    // Check for old v0 format
    const oldValue = figma.root.getPluginData('lastInvoiceNumber');
    if (oldValue && sequences.length === 0) {
      const defaultSeq = {
        id: generateId(),
        name: 'Invoice#',
        prefix: '',
        nextValue: oldValue,
        highestUsed: '0000',
        type: 'number',
        mode: 'compliance',
        locked: false,
        createdAt: Date.now()
      };
      saveSequences([defaultSeq]);
      setSelectedId(defaultSeq.id);
    }

    figma.root.setPluginData(MIGRATION_KEY, CURRENT_VERSION.toString());
  }
}

// ----------------------------------------------------------
// INCREMENT LOGIC
// ----------------------------------------------------------
function incrementNumber(value) {
  const num = parseInt(value, 10);
  return (num + 1).toString().padStart(value.length, '0');
}

function incrementLetter(value) {
  const chars = value.toUpperCase().split('');
  let i = chars.length - 1;

  while (i >= 0) {
    if (chars[i] === 'Z') {
      chars[i] = 'A';
      i--;
    } else {
      chars[i] = String.fromCharCode(chars[i].charCodeAt(0) + 1);
      return chars.join('');
    }
  }
  return 'A' + chars.join('');
}

function incrementValue(sequence) {
  if (sequence.type === 'letter') {
    return incrementLetter(sequence.nextValue);
  }
  return incrementNumber(sequence.nextValue);
}

function getFullValue(sequence) {
  return (sequence.prefix || '') + sequence.nextValue;
}

// Compare two values numerically (for compliance checks)
function compareValues(a, b, type) {
  if (type === 'letter') {
    return a.localeCompare(b);
  }
  return parseInt(a, 10) - parseInt(b, 10);
}

// ----------------------------------------------------------
// NODE HELPERS
// ----------------------------------------------------------
function getNodeLinkData(node) {
  if (node.type !== 'TEXT') return null;

  const sequenceId = node.getPluginData('sequenceId');
  const stampedValue = node.getPluginData('stampedValue');
  const stampedAt = node.getPluginData('stampedAt');

  if (!sequenceId) return null;

  return { sequenceId, stampedValue, stampedAt };
}

function setNodeLinkData(node, sequenceId, stampedValue) {
  node.setPluginData('sequenceId', sequenceId);
  node.setPluginData('stampedValue', stampedValue);
  node.setPluginData('stampedAt', Date.now().toString());
}

function clearNodeLinkData(node) {
  node.setPluginData('sequenceId', '');
  node.setPluginData('stampedValue', '');
  node.setPluginData('stampedAt', '');
}

// Check if a stamped value is duplicated (for detecting component duplicates)
function isStampedValueDuplicated(value, excludeNodeId) {
  let count = 0;

  function searchNode(node) {
    if (node.type === 'TEXT' && node.id !== excludeNodeId) {
      const stamped = node.getPluginData('stampedValue');
      if (stamped === value) {
        count++;
      }
    }
    if ('children' in node) {
      for (const child of node.children) {
        searchNode(child);
      }
    }
  }

  searchNode(figma.currentPage);
  return count > 0;
}

// Find all nodes linked to a sequence
function findNodesLinkedToSequence(sequenceId) {
  const nodes = [];

  function searchNode(node) {
    if (node.type === 'TEXT') {
      const linked = node.getPluginData('sequenceId');
      if (linked === sequenceId) {
        nodes.push(node);
      }
    }
    if ('children' in node) {
      for (const child of node.children) {
        searchNode(child);
      }
    }
  }

  searchNode(figma.currentPage);
  return nodes;
}

// ----------------------------------------------------------
// SELECTION ANALYSIS
// ----------------------------------------------------------
function analyzeSelection() {
  const selection = figma.currentPage.selection;

  if (selection.length !== 1) {
    return { state: 'none' };
  }

  const node = selection[0];

  if (node.type !== 'TEXT') {
    return { state: 'not-text' };
  }

  const linkData = getNodeLinkData(node);

  if (!linkData) {
    return {
      state: 'unlinked',
      nodeId: node.id,
      currentText: node.characters
    };
  }

  // Check if sequence still exists
  const sequences = getSequences();
  const sequence = sequences.find(s => s.id === linkData.sequenceId);

  if (!sequence) {
    return {
      state: 'broken-link',
      nodeId: node.id,
      stampedValue: linkData.stampedValue
    };
  }

  // Check if this is a duplicate (same stampedValue exists elsewhere)
  const isDuplicate = linkData.stampedValue &&
    isStampedValueDuplicated(linkData.stampedValue, node.id);

  if (isDuplicate || !linkData.stampedValue) {
    return {
      state: 'needs-stamp',
      nodeId: node.id,
      sequence: sequence,
      isDuplicate: isDuplicate
    };
  }

  return {
    state: 'stamped',
    nodeId: node.id,
    sequence: sequence,
    stampedValue: linkData.stampedValue
  };
}

function sendSelectionState() {
  const analysis = analyzeSelection();
  const sequences = getSequences();
  const selectedId = getSelectedId();
  const selectedSequence = sequences.find(s => s.id === selectedId);

  figma.ui.postMessage({
    type: 'selection-state',
    state: analysis.state,
    nodeId: analysis.nodeId,
    currentText: analysis.currentText,
    stampedValue: analysis.stampedValue,
    sequence: analysis.sequence,
    isDuplicate: analysis.isDuplicate,
    sequences: sequences,
    selectedId: selectedId,
    selectedSequence: selectedSequence
  });
}

// ----------------------------------------------------------
// SHOW UI
// ----------------------------------------------------------
figma.showUI(__html__, { width: 280, height: 400 });

// ----------------------------------------------------------
// INITIALIZATION
// ----------------------------------------------------------
function init() {
  try {
    migrateIfNeeded();
  } catch (e) {
    console.error('Migration error:', e);
  }
  try {
    sendSelectionState();
  } catch (e) {
    console.error('Selection state error:', e);
  }
}

init();

// Listen for selection changes
figma.on('selectionchange', () => {
  sendSelectionState();
});

// ----------------------------------------------------------
// MESSAGE HANDLERS
// ----------------------------------------------------------
figma.ui.onmessage = async (msg) => {

  // ----- SELECT SEQUENCE -----
  if (msg.type === 'select-sequence') {
    setSelectedId(msg.id);
    sendSelectionState();
  }

  // ----- CREATE SEQUENCE -----
  if (msg.type === 'create-sequence') {
    const sequences = getSequences();

    const newSeq = {
      id: generateId(),
      name: msg.name.trim(),
      prefix: (msg.prefix || '').trim(),
      nextValue: msg.startValue.trim(),
      highestUsed: '0', // Nothing used yet
      type: msg.sequenceType,
      mode: msg.mode, // 'compliance' or 'design'
      locked: false,
      createdAt: Date.now()
    };

    sequences.push(newSeq);
    saveSequences(sequences);
    setSelectedId(newSeq.id);

    figma.ui.postMessage({
      type: 'sequence-created',
      sequence: newSeq
    });

    sendSelectionState();
  }

  // ----- DELETE SEQUENCE -----
  if (msg.type === 'delete-sequence') {
    let sequences = getSequences();
    const sequence = sequences.find(s => s.id === msg.id);

    if (!sequence) return;

    // Compliance check: cannot delete if nodes are linked
    if (sequence.mode === 'compliance') {
      const linkedNodes = findNodesLinkedToSequence(msg.id);
      if (linkedNodes.length > 0) {
        figma.ui.postMessage({
          type: 'error',
          message: `Cannot delete: ${linkedNodes.length} document(s) are linked to this sequence`
        });
        return;
      }
    }

    sequences = sequences.filter(s => s.id !== msg.id);
    saveSequences(sequences);

    const newSelectedId = sequences.length > 0 ? sequences[0].id : null;
    setSelectedId(newSelectedId);

    figma.ui.postMessage({ type: 'sequence-deleted' });
    sendSelectionState();
  }

  // ----- LINK AND STAMP -----
  if (msg.type === 'link-and-stamp') {
    const selection = figma.currentPage.selection;

    if (selection.length !== 1 || selection[0].type !== 'TEXT') {
      figma.ui.postMessage({
        type: 'error',
        message: 'Select a single text layer'
      });
      return;
    }

    const node = selection[0];
    const sequences = getSequences();
    const sequence = sequences.find(s => s.id === msg.sequenceId);

    if (!sequence) {
      figma.ui.postMessage({ type: 'error', message: 'Sequence not found' });
      return;
    }

    // Get the next value to stamp
    const stampValue = getFullValue(sequence);

    // Load font and update text (handle mixed fonts and empty text)
    try {
      let fontName = node.fontName;
      if (fontName === figma.mixed) {
        if (node.characters.length > 0) {
          fontName = node.getRangeFontName(0, 1);
        } else {
          fontName = { family: "Inter", style: "Regular" };
        }
      }
      if (fontName === figma.mixed) {
        fontName = { family: "Inter", style: "Regular" };
      }
      await figma.loadFontAsync(fontName);
    } catch (e) {
      await figma.loadFontAsync({ family: "Inter", style: "Regular" });
    }
    node.characters = stampValue;

    // Store link data on the node
    setNodeLinkData(node, sequence.id, stampValue);

    // Update sequence: increment next, update highestUsed, lock if compliance
    sequence.nextValue = incrementValue(sequence);

    if (compareValues(stampValue.replace(sequence.prefix, ''), sequence.highestUsed, sequence.type) > 0) {
      sequence.highestUsed = stampValue.replace(sequence.prefix, '');
    }

    if (sequence.mode === 'compliance') {
      sequence.locked = true;
    }

    saveSequences(sequences);

    figma.ui.postMessage({
      type: 'stamped',
      stampedValue: stampValue,
      sequence: sequence
    });

    sendSelectionState();
  }

  // ----- STAMP (already linked, needs new number) -----
  if (msg.type === 'stamp') {
    const selection = figma.currentPage.selection;

    if (selection.length !== 1 || selection[0].type !== 'TEXT') {
      figma.ui.postMessage({ type: 'error', message: 'Select a text layer' });
      return;
    }

    const node = selection[0];
    const linkData = getNodeLinkData(node);

    if (!linkData) {
      figma.ui.postMessage({ type: 'error', message: 'Text is not linked' });
      return;
    }

    const sequences = getSequences();
    const sequence = sequences.find(s => s.id === linkData.sequenceId);

    if (!sequence) {
      figma.ui.postMessage({ type: 'error', message: 'Linked sequence not found' });
      return;
    }

    // Compliance check: cannot re-stamp if already has a unique value
    if (sequence.mode === 'compliance' && linkData.stampedValue) {
      const isDuplicate = isStampedValueDuplicated(linkData.stampedValue, node.id);
      if (!isDuplicate) {
        figma.ui.postMessage({
          type: 'error',
          message: 'Cannot re-stamp in compliance mode'
        });
        return;
      }
    }

    const stampValue = getFullValue(sequence);

    // Load font (handle mixed fonts and empty text)
    try {
      let fontName = node.fontName;
      if (fontName === figma.mixed) {
        if (node.characters.length > 0) {
          fontName = node.getRangeFontName(0, 1);
        } else {
          fontName = { family: "Inter", style: "Regular" };
        }
      }
      if (fontName === figma.mixed) {
        fontName = { family: "Inter", style: "Regular" };
      }
      await figma.loadFontAsync(fontName);
    } catch (e) {
      await figma.loadFontAsync({ family: "Inter", style: "Regular" });
    }
    node.characters = stampValue;

    setNodeLinkData(node, sequence.id, stampValue);

    sequence.nextValue = incrementValue(sequence);

    if (compareValues(stampValue.replace(sequence.prefix, ''), sequence.highestUsed, sequence.type) > 0) {
      sequence.highestUsed = stampValue.replace(sequence.prefix, '');
    }

    saveSequences(sequences);

    figma.ui.postMessage({
      type: 'stamped',
      stampedValue: stampValue,
      sequence: sequence
    });

    sendSelectionState();
  }

  // ----- UNLINK -----
  if (msg.type === 'unlink') {
    const selection = figma.currentPage.selection;

    if (selection.length !== 1 || selection[0].type !== 'TEXT') {
      return;
    }

    clearNodeLinkData(selection[0]);

    figma.ui.postMessage({ type: 'unlinked' });
    sendSelectionState();
  }

  // ----- RELINK (fix broken link) -----
  if (msg.type === 'relink') {
    const selection = figma.currentPage.selection;

    if (selection.length !== 1 || selection[0].type !== 'TEXT') {
      return;
    }

    const node = selection[0];
    node.setPluginData('sequenceId', msg.newSequenceId);
    // Keep existing stampedValue if any

    figma.ui.postMessage({ type: 'relinked' });
    sendSelectionState();
  }

  // ----- RESET (design mode only) -----
  if (msg.type === 'reset') {
    const sequences = getSequences();
    const sequence = sequences.find(s => s.id === msg.sequenceId);

    if (!sequence) return;

    // Compliance check
    if (sequence.mode === 'compliance') {
      const newNum = parseInt(msg.value, 10);
      const highestNum = parseInt(sequence.highestUsed, 10);

      if (newNum <= highestNum) {
        figma.ui.postMessage({
          type: 'error',
          message: `Cannot set below highest used (${sequence.prefix}${sequence.highestUsed})`
        });
        return;
      }
    }

    sequence.nextValue = msg.value.trim();
    saveSequences(sequences);

    figma.ui.postMessage({
      type: 'reset-done',
      sequence: sequence
    });

    sendSelectionState();
  }

  // ----- CLOSE -----
  if (msg.type === 'close') {
    figma.closePlugin();
  }
};
