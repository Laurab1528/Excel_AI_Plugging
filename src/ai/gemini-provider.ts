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
                this.model = this.genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });
      } catch (error) {
        console.error('Error inicializando Gemini:', error);
      }
    }
  }

  async isAvailable(): Promise<boolean> {
    if (!this.model) return false;
    
    // Simplemente verificar que el modelo esté inicializado
    // No hacer peticiones de prueba para no gastar créditos
    return true;
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

