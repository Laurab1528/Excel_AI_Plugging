import axios from 'axios';
import { AIProvider } from './types';

export class OllamaProvider implements AIProvider {
  name = 'Ollama';
  private baseUrl: string;
  private model: string;

  constructor(baseUrl: string = 'http://localhost:11434', model: string = 'llama3') {
    this.baseUrl = baseUrl;
    this.model = model;
  }

  async isAvailable(): Promise<boolean> {
    try {
      const response = await axios.get(`${this.baseUrl}/api/tags`, {
        timeout: 2000
      });
      
      // Verificar si el modelo está disponible
      const models = response.data.models || [];
      return models.some((m: any) => m.name.includes(this.model));
    } catch (error) {
      console.log('Ollama no está disponible o no tiene el modelo instalado');
      return false;
    }
  }

  async generateResponse(prompt: string, context?: string): Promise<string> {
    try {
      const fullPrompt = context 
        ? `Contexto: ${context}\n\nInstrucción: ${prompt}` 
        : prompt;

      const response = await axios.post(
        `${this.baseUrl}/api/generate`,
        {
          model: this.model,
          prompt: fullPrompt,
          stream: false
        },
        {
          timeout: 60000 // 60 segundos para respuestas largas
        }
      );

      return response.data.response || '';
    } catch (error: any) {
      if (error.code === 'ECONNREFUSED') {
        throw new Error('Ollama no está ejecutándose. Inicia Ollama primero.');
      }
      throw error;
    }
  }
}

