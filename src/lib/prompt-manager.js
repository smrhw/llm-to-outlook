import { storage } from './storage.js';

/**
 * Generate a simple UUID
 * @returns {string}
 */
function generateId() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0;
        const v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

export const promptManager = {
    /**
     * Get all saved prompts
     * @returns {Array<{id: string, name: string, instruction: string}>}
     */
    getAll() {
        return storage.getSavedPrompts();
    },

    /**
     * Get a prompt by ID
     * @param {string} id 
     * @returns {Object|null}
     */
    getById(id) {
        const prompts = this.getAll();
        return prompts.find(p => p.id === id) || null;
    },

    /**
     * Add a new prompt
     * @param {string} name 
     * @param {string} instruction 
     * @returns {Object} The created prompt
     */
    add(name, instruction) {
        const prompts = this.getAll();
        const newPrompt = {
            id: generateId(),
            name: name.trim(),
            instruction: instruction.trim()
        };
        prompts.push(newPrompt);
        storage.setSavedPrompts(prompts);
        return newPrompt;
    },

    /**
     * Update an existing prompt
     * @param {string} id 
     * @param {string} name 
     * @param {string} instruction 
     * @returns {boolean}
     */
    update(id, name, instruction) {
        const prompts = this.getAll();
        const index = prompts.findIndex(p => p.id === id);
        if (index === -1) return false;

        prompts[index] = {
            ...prompts[index],
            name: name.trim(),
            instruction: instruction.trim()
        };
        storage.setSavedPrompts(prompts);
        return true;
    },

    /**
     * Delete a prompt
     * @param {string} id 
     * @returns {boolean}
     */
    delete(id) {
        const prompts = this.getAll();
        const filtered = prompts.filter(p => p.id !== id);
        if (filtered.length === prompts.length) return false;

        storage.setSavedPrompts(filtered);
        return true;
    },

    /**
     * Add some default prompts if none exist
     */
    initDefaults() {
        if (this.getAll().length === 0) {
            this.add('Fix Grammar', 'Fix any grammatical errors in this text while maintaining the original meaning and tone.');
            this.add('Make Professional', 'Rewrite this text to be more professional and formal while keeping the same message.');
            this.add('Summarize', 'Provide a brief, concise summary of this text.');
            this.add('Translate to French', 'Translate this text to French.');
            this.add('Make Friendly', 'Rewrite this text to be warmer and more friendly while keeping the same message.');
        }
    }
};
