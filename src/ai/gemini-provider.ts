import { GoogleGenerativeAI } from '@google/generative-ai';
import { AIProvider } from './types';

export class GeminiProvider implements AIProvider {
  name = 'Gemini';
  private genAI: GoogleGenerativeAI | null = null;
  private model: any = null;
  private apiKey: string;

  constructor(apiKey: string) {
    this.apiKey = apiKey;
    if (apiKey && apiKey !== 'tu_api_key_aqui') {
      try {
        this.genAI = new GoogleGenerativeAI(apiKey);
        this.model = this.genAI.getGenerativeModel({ model: 'gemini-pro' });
      } catch (error) {
        console.error('Error inicializando Gemini:', error);
      }
    }
  }

  async isAvailable(): Promise<boolean> {
    if (!this.model) return false;
    
    try {
      // Test simple para verificar que la API funciona
      const result = await this.model.generateContent('test');
      return !!result;
    } catch (error: any) {
      // Verificar si es un error de cuota
      if (error?.message?.includes('quota') || error?.message?.includes('limit')) {
        console.warn('Gemini: Cuota agotada');
        return false;
      }
      console.error('Gemini no disponible:', error);
      return false;
    }
  }

  async generateResponse(prompt: string, context?: string): Promise<string> {
    if (!this.model) {
      throw new Error('Gemini no está configurado. Verifica tu API key.');
    }

    try {
      const fullPrompt = context 
        ? `Contexto: ${context}\n\nInstrucción: ${prompt}` 
        : prompt;

      const result = await this.model.generateContent(fullPrompt);
      const response = await result.response;
      return response.text();
    } catch (error: any) {
      if (error?.message?.includes('quota') || error?.message?.includes('limit')) {
        throw new Error('QUOTA_EXCEEDED');
      }
      throw error;
    }
  }
}

