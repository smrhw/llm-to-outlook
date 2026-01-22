import { storage } from './storage.js';

/**
 * System prompt for structured output
 */
const SYSTEM_PROMPT = `You are an expert email assistant. When given an email body and an instruction, you must:
1. Generate an appropriate email subject line
2. Generate the improved/modified email body based on the instruction

IMPORTANT: You MUST respond with valid JSON in this exact format:
{
  "subject": "Your suggested subject line here",
  "body": "<p>Your HTML-formatted email body here</p>"
}

For the body, use RICH HTML FORMATTING to create visually appealing, well-structured emails:
- Use <p> tags for paragraphs with proper spacing between ideas
- Use <strong> for bold emphasis on important words or phrases
- Use <em> for italic text when appropriate
- Use <ul> and <li> for bullet point lists (great for action items, key points)
- Use <ol> and <li> for numbered/ordered lists (great for steps, priorities)
- Use <br> for line breaks within paragraphs when needed
- Use emojis SPARINGLY and only when they naturally add warmth or clarity:
  • Limit usage to 1-2 relevant emojis maximum, and only if the tone is friendly or informal
  • Examples where they may be appropriate: 👋 for greetings, ✅ for confirmations, or 🎉 for celebrations
  • Avoid emojis in formal or strictly professional correspondence

OBJECT PRESERVATION:
- The input may contain placeholders like [[TABLE_1]] or [[IMAGE_1]].
- These represent embedded tables or images that MUST be preserved exactly as-is.
- Include these placeholders in your output in their appropriate relative positions.
- Do NOT modify, remove, or rewrite the placeholder text.

- Keep the formatting clean, professional, and visually structured
- Do not include any text outside the JSON object
- Do not use markdown syntax - use only HTML tags`;


/**
 * Base LLM Provider class
 */
class LLMProvider {
    constructor(config) {
        this.apiKey = config.apiKey;
        this.endpoint = config.endpoint;
        this.model = config.model;
    }

    async processText(selectedText, instruction) {
        throw new Error('processText must be implemented by subclass');
    }
}

/**
 * OpenAI Provider (also works with OpenAI-compatible APIs)
 */
class OpenAIProvider extends LLMProvider {
    async processText(emailBody, instruction) {
        const response = await fetch(`${this.endpoint}/chat/completions`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${this.apiKey}`
            },
            body: JSON.stringify({
                model: this.model,
                messages: [
                    {
                        role: 'system',
                        content: SYSTEM_PROMPT
                    },
                    {
                        role: 'user',
                        content: `Instruction: ${instruction}\n\nEmail content:\n${emailBody}`
                    }
                ],
                temperature: 0.7,
                response_format: { type: "json_object" }
            })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `OpenAI API error: ${response.status}`);
        }

        const data = await response.json();
        return data.choices[0]?.message?.content || '{}';
    }
}

/**
 * Claude (Anthropic) Provider
 */
class ClaudeProvider extends LLMProvider {
    async processText(emailBody, instruction) {
        const response = await fetch(`${this.endpoint}/v1/messages`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': this.apiKey,
                'anthropic-version': '2023-06-01',
                'anthropic-dangerous-direct-browser-access': 'true'
            },
            body: JSON.stringify({
                model: this.model,
                max_tokens: 4096,
                system: SYSTEM_PROMPT,
                messages: [
                    {
                        role: 'user',
                        content: `Instruction: ${instruction}\n\nEmail content:\n${emailBody}`
                    }
                ]
            })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `Claude API error: ${response.status}`);
        }

        const data = await response.json();
        return data.content[0]?.text || '{}';
    }
}

/**
 * Google Gemini Provider
 */
class GeminiProvider extends LLMProvider {
    async processText(emailBody, instruction) {
        const url = `${this.endpoint}/v1beta/models/${this.model}:generateContent?key=${this.apiKey}`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                contents: [
                    {
                        parts: [
                            {
                                text: `${SYSTEM_PROMPT}\n\nInstruction: ${instruction}\n\nEmail content:\n${emailBody}`
                            }
                        ]
                    }
                ],
                generationConfig: {
                    temperature: 0.7,
                    responseMimeType: "application/json"
                }
            })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `Gemini API error: ${response.status}`);
        }

        const data = await response.json();
        return data.candidates[0]?.content?.parts[0]?.text || '{}';
    }
}

/**
 * Get the appropriate provider instance based on settings
 * @returns {LLMProvider}
 */
export function getProvider() {
    const settings = storage.getProviderSettings();
    const activeProvider = settings.activeProvider;
    const config = settings.providers[activeProvider];

    if (!config.apiKey) {
        throw new Error(`No API key configured for ${activeProvider}. Please configure in settings.`);
    }

    switch (activeProvider) {
        case 'openai':
        case 'custom':
            return new OpenAIProvider(config);
        case 'claude':
            return new ClaudeProvider(config);
        case 'gemini':
            return new GeminiProvider(config);
        default:
            throw new Error(`Unknown provider: ${activeProvider}`);
    }
}

/**
 * Process text using the configured LLM provider
 * @param {string} emailBody 
 * @param {string} instruction 
 * @returns {Promise<string>}
 */
export async function processText(emailBody, instruction) {
    const provider = getProvider();
    return provider.processText(emailBody, instruction);
}
