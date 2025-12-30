/**
 * Storage utility for managing localStorage with prefixed keys
 */
const STORAGE_PREFIX = 'outlook_ai_assistant_';

export const storage = {
  /**
   * Get a value from localStorage
   * @param {string} key 
   * @returns {any}
   */
  get(key) {
    try {
      const value = localStorage.getItem(STORAGE_PREFIX + key);
      return value ? JSON.parse(value) : null;
    } catch (e) {
      console.error('Storage get error:', e);
      return null;
    }
  },

  /**
   * Set a value in localStorage
   * @param {string} key 
   * @param {any} value 
   */
  set(key, value) {
    try {
      localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(value));
    } catch (e) {
      console.error('Storage set error:', e);
    }
  },

  /**
   * Remove a value from localStorage
   * @param {string} key 
   */
  remove(key) {
    try {
      localStorage.removeItem(STORAGE_PREFIX + key);
    } catch (e) {
      console.error('Storage remove error:', e);
    }
  },

  /**
   * Get LLM provider settings
   * @returns {Object}
   */
  getProviderSettings() {
    return this.get('provider_settings') || {
      activeProvider: 'openai',
      providers: {
        openai: {
          apiKey: '',
          endpoint: 'https://api.openai.com/v1',
          model: 'gpt-4o-mini'
        },
        claude: {
          apiKey: '',
          endpoint: 'https://api.anthropic.com',
          model: 'claude-3-5-sonnet-20241022'
        },
        gemini: {
          apiKey: '',
          endpoint: 'https://generativelanguage.googleapis.com',
          model: 'gemini-1.5-flash'
        },
        custom: {
          apiKey: '',
          endpoint: '',
          model: ''
        }
      }
    };
  },

  /**
   * Save LLM provider settings
   * @param {Object} settings 
   */
  setProviderSettings(settings) {
    this.set('provider_settings', settings);
  },

  /**
   * Get saved prompts
   * @returns {Array}
   */
  getSavedPrompts() {
    return this.get('saved_prompts') || [];
  },

  /**
   * Save prompts array
   * @param {Array} prompts 
   */
  setSavedPrompts(prompts) {
    this.set('saved_prompts', prompts);
  }
};
