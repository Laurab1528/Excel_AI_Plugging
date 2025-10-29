import { AIProvider, AIConfig, AIProviderType } from './types';
import { GeminiProvider } from './gemini-provider';
import { OllamaProvider } from './ollama-provider';

export class AIManager {
  private providers: Map<AIProviderType, AIProvider> = new Map();
  private currentProvider: AIProvider | null = null;
  private fallbackAttempted = false;

  constructor(config: AIConfig) {
    // Inicializar Gemini si hay API key
    if (config.geminiApiKey) {
      this.providers.set(
        AIProviderType.GEMINI,
        new GeminiProvider(config.geminiApiKey)
      );
    }

    // Siempre inicializar Ollama como fallback
    this.providers.set(
      AIProviderType.OLLAMA,
      new OllamaProvider(config.ollamaUrl, config.ollamaModel)
    );
  }

  async initialize(): Promise<void> {
    // Intentar usar Gemini primero
    const gemini = this.providers.get(AIProviderType.GEMINI);
    if (gemini && await gemini.isAvailable()) {
      this.currentProvider = gemini;
      console.log('✅ Usando Gemini como proveedor de IA');
      return;
    }

    // Si Gemini no está disponible, usar Ollama
    const ollama = this.providers.get(AIProviderType.OLLAMA);
    if (ollama && await ollama.isAvailable()) {
      this.currentProvider = ollama;
      console.log('✅ Usando Ollama como proveedor de IA');
      return;
    }

    throw new Error('No hay ningún proveedor de IA disponible. Configura Gemini API key o inicia Ollama.');
  }

  async generateResponse(prompt: string, context?: string): Promise<{
    response: string;
    provider: string;
    fallbackUsed: boolean;
  }> {
    if (!this.currentProvider) {
      await this.initialize();
    }

    try {
      const response = await this.currentProvider!.generateResponse(prompt, context);
      this.fallbackAttempted = false; // Reset flag on success
      
      return {
        response,
        provider: this.currentProvider!.name,
        fallbackUsed: false
      };
    } catch (error: any) {
      // Si Gemini falla por cuota, intentar Ollama
      if (error.message === 'QUOTA_EXCEEDED' && !this.fallbackAttempted) {
        console.warn('⚠️ Cuota de Gemini agotada, cambiando a Ollama...');
        this.fallbackAttempted = true;
        
        const ollama = this.providers.get(AIProviderType.OLLAMA);
        if (ollama && await ollama.isAvailable()) {
          this.currentProvider = ollama;
          const response = await ollama.generateResponse(prompt, context);
          
          return {
            response,
            provider: ollama.name,
            fallbackUsed: true
          };
        }
      }

      throw error;
    }
  }

  getCurrentProvider(): string {
    return this.currentProvider?.name || 'Ninguno';
  }

  async switchProvider(type: AIProviderType): Promise<boolean> {
    const provider = this.providers.get(type);
    if (!provider) {
      return false;
    }

    if (await provider.isAvailable()) {
      this.currentProvider = provider;
      console.log(`Cambiado a proveedor: ${provider.name}`);
      return true;
    }

    return false;
  }
}

