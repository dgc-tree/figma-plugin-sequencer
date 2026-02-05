// ============================================================
// SEQUENCER - Figma Plugin
// ============================================================
// This plugin helps you manage multiple sequences of numbers
// or letters across your Figma documents. Perfect for invoices,
// estimates, revisions, versions, and more.
//
// Features:
// - Multiple named sequences (Invoice#, Estimate#, PO#, etc.)
// - Number sequences (5100 → 5101 → 5102)
// - Letter sequences (A → B → ... → Z → AA → AB)
// - Optional prefix (Q0001, INV-100, etc.)
// - Batch updates all matching text layers
// - Persists between sessions
// ============================================================

// ----------------------------------------------------------
// STORAGE KEYS
// ----------------------------------------------------------
const SEQUENCES_KEY = 'sequences';        // JSON array of all sequences
const SELECTED_KEY = 'selectedSequence';  // ID of currently selected sequence
const MIGRATION_KEY = 'migrationVersion'; // Track data format version
const CURRENT_VERSION = 1;

// ----------------------------------------------------------
// STORAGE HELPERS
// ----------------------------------------------------------

// Get all sequences from plugin data
function getSequences() {
  const data = figma.root.getPluginData(SEQUENCES_KEY);
  return data ? JSON.parse(data) : [];
}

// Save all sequences to plugin data
function saveSequences(sequences) {
  figma.root.setPluginData(SEQUENCES_KEY, JSON.stringify(sequences));
}

// Get the ID of the currently selected sequence
function getSelectedId() {
  return figma.root.getPluginData(SELECTED_KEY) || null;
}

// Set the currently selected sequence
function setSelectedId(id) {
  figma.root.setPluginData(SELECTED_KEY, id || '');
}

// Generate a unique ID for new sequences
function generateId() {
  return 'seq_' + Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

// ----------------------------------------------------------
// MIGRATION
// Handles upgrading from old single-number format
// ----------------------------------------------------------
function migrateIfNeeded() {
  const version = figma.root.getPluginData(MIGRATION_KEY);

  if (!version || parseInt(version) < CURRENT_VERSION) {
    // Check for old single-number format
    const oldValue = figma.root.getPluginData('lastInvoiceNumber');

    if (oldValue) {
      // Migrate: Create a default sequence from old data
      const defaultSequence = {
        id: generateId(),
        name: 'Invoice#',
        prefix: '',
        value: oldValue,
        type: 'number'
      };

      saveSequences([defaultSequence]);
      setSelectedId(defaultSequence.id);

      // Note: We keep the old key for rollback safety
      console.log('Migrated from old format:', defaultSequence);
    }

    // Mark migration complete
    figma.root.setPluginData(MIGRATION_KEY, CURRENT_VERSION.toString());
  }
}

// ----------------------------------------------------------
// INCREMENT LOGIC
// ----------------------------------------------------------

// Increment a number sequence, preserving leading zeros
// "0099" → "0100", "5100" → "5101"
function incrementNumber(value) {
  const num = parseInt(value, 10);
  const nextNum = num + 1;
  // Preserve the original length with leading zeros
  return nextNum.toString().padStart(value.length, '0');
}

// Increment a letter sequence (Excel column style)
// A → B, Z → AA, AZ → BA, ZZ → AAA
function incrementLetter(value) {
  const chars = value.toUpperCase().split('');
  let i = chars.length - 1;

  while (i >= 0) {
    if (chars[i] === 'Z') {
      chars[i] = 'A'; // Roll over Z → A
      i--; // Carry to next position
    } else {
      // Increment this character and we're done
      chars[i] = String.fromCharCode(chars[i].charCodeAt(0) + 1);
      return chars.join('');
    }
  }

  // All positions were Z, add new 'A' at start: ZZ → AAA
  return 'A' + chars.join('');
}

// Unified increment function based on sequence type
function incrementValue(sequence) {
  if (sequence.type === 'letter') {
    return incrementLetter(sequence.value);
  } else {
    return incrementNumber(sequence.value);
  }
}

// Get the full formatted value: prefix + value
// Example: prefix="Q", value="0001" → "Q0001"
function getFullValue(sequence) {
  const prefix = sequence.prefix || '';
  return prefix + sequence.value;
}

// Get the next full formatted value
function getNextFullValue(sequence) {
  const prefix = sequence.prefix || '';
  return prefix + incrementValue(sequence);
}

// ----------------------------------------------------------
// TEXT NODE SEARCH
// Finds all text layers matching a specific value
// ----------------------------------------------------------
function findTextNodesWithValue(value) {
  const matches = [];
  const searchString = value.toString().trim();

  function searchNode(node) {
    if (node.type === 'TEXT') {
      const text = node.characters.trim();
      // Case-insensitive match for letter sequences
      if (text.toLowerCase() === searchString.toLowerCase()) {
        matches.push(node);
      }
    }

    // Recursively search children (frames, groups, components, etc.)
    if ('children' in node) {
      for (const child of node.children) {
        searchNode(child);
      }
    }
  }

  searchNode(figma.currentPage);
  return matches;
}

// ----------------------------------------------------------
// SHOW THE UI
// ----------------------------------------------------------
figma.showUI(__html__, { width: 260, height: 340 });

// ----------------------------------------------------------
// INITIALIZATION
// ----------------------------------------------------------
function init() {
  // Run migration if needed (handles old format)
  migrateIfNeeded();

  const sequences = getSequences();
  let selectedId = getSelectedId();

  // Validate selected sequence exists
  let selectedSequence = null;
  if (selectedId) {
    selectedSequence = sequences.find(s => s.id === selectedId);
  }

  // If no valid selection, default to first sequence
  if (!selectedSequence && sequences.length > 0) {
    selectedSequence = sequences[0];
    setSelectedId(selectedSequence.id);
    selectedId = selectedSequence.id;
  }

  // Send initial state to UI
  figma.ui.postMessage({
    type: 'init',
    sequences: sequences,
    selectedId: selectedId,
    selectedSequence: selectedSequence
  });
}

// Run initialization
init();

// ----------------------------------------------------------
// MESSAGE HANDLERS
// Responds to messages from the UI
// ----------------------------------------------------------
figma.ui.onmessage = async (msg) => {

  // ----- SELECT: User chose a different sequence -----
  if (msg.type === 'select-sequence') {
    setSelectedId(msg.id);
    const sequences = getSequences();
    const selected = sequences.find(s => s.id === msg.id);

    figma.ui.postMessage({
      type: 'sequence-selected',
      sequence: selected
    });
  }

  // ----- CREATE: User creates a new sequence -----
  if (msg.type === 'create-sequence') {
    const sequences = getSequences();

    const newSeq = {
      id: generateId(),
      name: msg.name.trim(),
      prefix: (msg.prefix || '').trim(),
      value: msg.value.trim(),
      type: msg.sequenceType // 'number' or 'letter'
    };

    sequences.push(newSeq);
    saveSequences(sequences);
    setSelectedId(newSeq.id);

    figma.ui.postMessage({
      type: 'sequence-created',
      sequences: sequences,
      selectedId: newSeq.id,
      selectedSequence: newSeq
    });
  }

  // ----- DELETE: User deletes a sequence -----
  if (msg.type === 'delete-sequence') {
    let sequences = getSequences();
    sequences = sequences.filter(s => s.id !== msg.id);
    saveSequences(sequences);

    // Select another sequence or clear selection
    let newSelectedId = null;
    let newSelected = null;

    if (sequences.length > 0) {
      newSelected = sequences[0];
      newSelectedId = newSelected.id;
    }

    setSelectedId(newSelectedId);

    figma.ui.postMessage({
      type: 'sequence-deleted',
      sequences: sequences,
      selectedId: newSelectedId,
      selectedSequence: newSelected
    });
  }

  // ----- UPDATE: Increment and update all matching text layers -----
  if (msg.type === 'update') {
    const sequences = getSequences();
    const sequence = sequences.find(s => s.id === msg.sequenceId);

    if (!sequence) {
      figma.ui.postMessage({ type: 'error', message: 'Sequence not found' });
      return;
    }

    // Get full formatted values (with prefix/suffix)
    const currentFullValue = getFullValue(sequence);
    const nextFullValue = getNextFullValue(sequence);

    // Find all text layers with the current full value
    const matchingNodes = findTextNodesWithValue(currentFullValue);

    if (matchingNodes.length === 0) {
      figma.ui.postMessage({
        type: 'info',
        message: `No text layers found with "${currentFullValue}"`
      });
      return;
    }

    // Update each matching text layer
    for (const node of matchingNodes) {
      // Must load font before changing text (Figma API requirement)
      await figma.loadFontAsync(node.fontName);
      node.characters = nextFullValue;
    }

    // Update and save the sequence (only the value part increments)
    sequence.value = incrementValue(sequence);
    saveSequences(sequences);

    figma.ui.postMessage({
      type: 'updated',
      sequence: sequence,
      count: matchingNodes.length
    });
  }

  // ----- RESET: Set sequence to a specific value -----
  if (msg.type === 'reset') {
    const sequences = getSequences();
    const sequence = sequences.find(s => s.id === msg.sequenceId);

    if (sequence) {
      sequence.value = msg.value.trim();
      saveSequences(sequences);

      figma.ui.postMessage({
        type: 'sequence-reset',
        sequence: sequence
      });
    }
  }

  // ----- CLOSE: User closed the plugin -----
  if (msg.type === 'close') {
    figma.closePlugin();
  }
};
