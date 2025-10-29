// Tipos para el sistema de IA
export interface AIProvider {
  name: string;
  generateResponse(prompt: string, context?: string): Promise<string>;
  isAvailable(): Promise<boolean>;
}

export interface AIConfig {
  geminiApiKey?: string;
  ollamaUrl?: string;
  ollamaModel?: string;
}

export enum AIProviderType {
  GEMINI = 'gemini',
  OLLAMA = 'ollama'
}

export interface ChatMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
}

