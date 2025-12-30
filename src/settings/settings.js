import { storage } from '../lib/storage.js';
import { promptManager } from '../lib/prompt-manager.js';

// DOM Elements
let elements = {};
let currentEditingPromptId = null;

// Provider default values
const providerDefaults = {
    openai: {
        endpoint: 'https://api.openai.com/v1',
        model: 'gpt-4o-mini',
        endpointHint: 'Default: https://api.openai.com/v1',
        modelHint: 'e.g., gpt-4o, gpt-4o-mini, gpt-3.5-turbo'
    },
    claude: {
        endpoint: 'https://api.anthropic.com',
        model: 'claude-3-5-sonnet-20241022',
        endpointHint: 'Default: https://api.anthropic.com',
        modelHint: 'e.g., claude-3-5-sonnet-20241022, claude-3-haiku-20240307'
    },
    gemini: {
        endpoint: 'https://generativelanguage.googleapis.com',
        model: 'gemini-1.5-flash',
        endpointHint: 'Default: https://generativelanguage.googleapis.com',
        modelHint: 'e.g., gemini-1.5-flash, gemini-1.5-pro, gemini-pro'
    },
    custom: {
        endpoint: '',
        model: '',
        endpointHint: 'OpenAI-compatible API endpoint URL',
        modelHint: 'Model name as required by the API'
    }
};

/**
 * Initialize when Office is ready or DOM is loaded
 * We use both to be safe
 */
function onReady() {
    initializeApp();
}

if (typeof Office !== 'undefined') {
    Office.onReady(onReady);
} else {
    document.addEventListener('DOMContentLoaded', onReady);
}

/**
 * Initialize the settings page
 */
function initializeApp() {
    console.log('Initializing settings app...');

    // Ensure default prompts exist if storage is empty
    promptManager.initDefaults();

    // Initial render
    renderPromptsList();
    loadSettings();

    // Set up event listeners
    setupEventListeners();
}

/**
 * Common selector utility
 */
const $ = (id) => document.getElementById(id);

/**
 * Set up event listeners
 */
function setupEventListeners() {
    // Static button listeners - more reliable than document delegation for core actions
    const backBtn = $('backBtn');
    if (backBtn) backBtn.onclick = () => window.location.href = '../taskpane/taskpane.html';

    const addBtn = $('addPromptBtn');
    if (addBtn) addBtn.onclick = () => openPromptModal();

    const saveSettingsBtn = $('saveSettingsBtn');
    if (saveSettingsBtn) saveSettingsBtn.onclick = () => saveSettings();

    const savePromptBtn = $('savePromptBtn');
    if (savePromptBtn) savePromptBtn.onclick = () => savePrompt();

    const closeModalBtn = $('closeModalBtn');
    if (closeModalBtn) closeModalBtn.onclick = () => closePromptModal();

    const cancelModalBtn = $('cancelModalBtn');
    if (cancelModalBtn) cancelModalBtn.onclick = () => closePromptModal();

    const toggleApiKeyBtn = $('toggleApiKey');
    if (toggleApiKeyBtn) toggleApiKeyBtn.onclick = () => toggleApiKeyVisibility();

    // Universal delegator for dynamic items (Edit/Delete buttons)
    document.addEventListener('click', (e) => {
        const target = e.target;

        // Delete Prompt
        const deleteBtn = target.closest('.delete-prompt-btn');
        if (deleteBtn) {
            const item = deleteBtn.closest('.prompt-item');
            if (item && item.dataset.id) {
                deletePrompt(item.dataset.id);
            }
            return;
        }

        // Edit Prompt
        const editBtn = target.closest('.edit-prompt-btn');
        if (editBtn) {
            const item = editBtn.closest('.prompt-item');
            if (item && item.dataset.id) {
                openPromptModal(item.dataset.id);
            }
            return;
        }

        // Backdrop click to close modal
        if (target.classList.contains('modal-backdrop')) {
            closePromptModal();
        }
    });

    // Provider selection radios
    document.querySelectorAll('input[name="provider"]').forEach(radio => {
        radio.addEventListener('change', onProviderChange);
    });
}

/**
 * Load settings from storage
 */
function loadSettings() {
    try {
        const settings = storage.getProviderSettings();
        const activeProvider = settings.activeProvider;

        const radio = document.querySelector(`input[name="provider"][value="${activeProvider}"]`);
        if (radio) radio.checked = true;

        loadProviderConfig(activeProvider, settings.providers[activeProvider]);
    } catch (e) {
        console.error('Error loading settings:', e);
    }
}

/**
 * Load provider-specific configuration
 */
function loadProviderConfig(provider, config) {
    const defaults = providerDefaults[provider];
    if (!defaults) return;

    const apiKeyEl = $('apiKey');
    const endpointEl = $('endpoint');
    const modelEl = $('model');

    if (apiKeyEl) apiKeyEl.value = config?.apiKey || '';
    if (endpointEl) endpointEl.value = config?.endpoint || defaults.endpoint;
    if (modelEl) modelEl.value = config?.model || defaults.model;

    const endpointHintEl = $('endpointHint');
    const modelHintEl = $('modelHint');
    if (endpointHintEl) endpointHintEl.textContent = defaults.endpointHint;
    if (modelHintEl) modelHintEl.textContent = defaults.modelHint;
}

/**
 * Handle provider change
 */
function onProviderChange(e) {
    const provider = e.target.value;
    const settings = storage.getProviderSettings();
    const config = settings.providers[provider];
    loadProviderConfig(provider, config);
}

/**
 * Toggle API key visibility
 */
function toggleApiKeyVisibility() {
    const apiKeyEl = $('apiKey');
    const toggleBtn = $('toggleApiKey');
    if (!apiKeyEl || !toggleBtn) return;

    const eyeOpen = toggleBtn.querySelector('.eye-open');
    const eyeClosed = toggleBtn.querySelector('.eye-closed');

    if (apiKeyEl.type === 'password') {
        apiKeyEl.type = 'text';
        if (eyeOpen) eyeOpen.style.display = 'none';
        if (eyeClosed) eyeClosed.style.display = 'block';
    } else {
        apiKeyEl.type = 'password';
        if (eyeOpen) eyeOpen.style.display = 'block';
        if (eyeClosed) eyeClosed.style.display = 'none';
    }
}

/**
 * Save settings to storage
 */
function saveSettings() {
    try {
        const checkedRadio = document.querySelector('input[name="provider"]:checked');
        if (!checkedRadio) return;

        const selectedProvider = checkedRadio.value;
        const settings = storage.getProviderSettings();

        settings.activeProvider = selectedProvider;
        settings.providers[selectedProvider] = {
            apiKey: ($('apiKey')?.value || '').trim(),
            endpoint: ($('endpoint')?.value || '').trim(),
            model: ($('model')?.value || '').trim()
        };

        storage.setProviderSettings(settings);
        showToast('Settings saved successfully!', 'success');
    } catch (e) {
        showToast('Error saving settings: ' + e.message, 'error');
    }
}

/**
 * Render the prompts list
 */
function renderPromptsList() {
    const promptsListEl = $('promptsList');
    if (!promptsListEl) return;

    const prompts = promptManager.getAll();

    if (prompts.length === 0) {
        promptsListEl.innerHTML = `
            <div class="empty-state">
                <p>No saved prompts yet.</p>
                <p>Click "Add" to create your first prompt.</p>
            </div>
        `;
        return;
    }

    promptsListEl.innerHTML = prompts.map(prompt => `
        <div class="prompt-item" data-id="${prompt.id}">
            <div class="prompt-info">
                <div class="prompt-name">${escapeHtml(prompt.name)}</div>
                <div class="prompt-instruction">${escapeHtml(prompt.instruction)}</div>
            </div>
            <div class="prompt-actions">
                <button class="btn btn-secondary edit-prompt-btn" title="Edit">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="pointer-events: none;">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                    </svg>
                </button>
                <button class="btn btn-danger delete-prompt-btn" title="Delete">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="pointer-events: none;">
                        <polyline points="3 6 5 6 21 6"></polyline>
                        <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                    </svg>
                </button>
            </div>
        </div>
    `).join('');
}

/**
 * Modal state management
 */
function openPromptModal(promptId = null) {
    const modalEl = $('promptModal');
    const modalTitleEl = $('modalTitle');
    const nameEl = $('promptName');
    const instructionEl = $('promptInstruction');

    if (!modalEl || !modalTitleEl || !nameEl || !instructionEl) return;

    currentEditingPromptId = promptId;

    if (promptId) {
        const prompt = promptManager.getById(promptId);
        if (prompt) {
            modalTitleEl.textContent = 'Edit Prompt';
            nameEl.value = prompt.name;
            instructionEl.value = prompt.instruction;
        }
    } else {
        modalTitleEl.textContent = 'Add Prompt';
        nameEl.value = '';
        instructionEl.value = '';
    }

    modalEl.style.display = 'flex';
    nameEl.focus();
}

/**
 * Close modal
 */
function closePromptModal() {
    const modalEl = $('promptModal');
    if (modalEl) modalEl.style.display = 'none';
    currentEditingPromptId = null;
}

/**
 * Save prompt
 */
function savePrompt() {
    const nameInput = $('promptName');
    const instructionInput = $('promptInstruction');

    if (!nameInput || !instructionInput) return;

    const name = nameInput.value.trim();
    const instruction = instructionInput.value.trim();

    if (!name || !instruction) {
        showToast('Please fill in both name and instruction', 'error');
        return;
    }

    try {
        if (currentEditingPromptId) {
            promptManager.update(currentEditingPromptId, name, instruction);
            showToast('Prompt updated!', 'success');
        } else {
            promptManager.add(name, instruction);
            showToast('Prompt saved successfully!', 'success');
        }

        closePromptModal();
        renderPromptsList();
    } catch (e) {
        console.error('Error saving prompt:', e);
        showToast('Error: ' + e.message, 'error');
    }
}

/**
 * Delete prompt
 */
function deletePrompt(id) {
    // Some Outlook environments block native confirm()
    // For now, we'll just delete to ensure it works, or use a custom modal later
    try {
        const deleted = promptManager.delete(id);
        if (deleted) {
            showToast('Prompt deleted', 'success');
            renderPromptsList();
        }
    } catch (e) {
        showToast('Error deleting: ' + e.message, 'error');
    }
}

/**
 * Toast helper
 */
function showToast(message, type = 'success') {
    const toastEl = $('toast');
    if (!toastEl) return;

    toastEl.textContent = message;
    toastEl.className = `toast ${type}`;
    toastEl.style.display = 'block';

    setTimeout(() => {
        toastEl.style.display = 'none';
    }, 3000);
}

/**
 * Simple HTML escaper
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text || '';
    return div.innerHTML;
}

