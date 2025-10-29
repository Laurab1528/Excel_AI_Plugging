import * as React from 'react';
import './App.css';
import { AIManager } from '../ai/ai-manager';
import { ExcelService } from '../services/excel-service';
import { WordService } from '../services/word-service';
import { PowerPointService } from '../services/powerpoint-service';

/* global Office */

interface AppProps {
  host: Office.HostType;
}

const App: React.FC<AppProps> = ({ host }) => {
  const [prompt, setPrompt] = React.useState('');
  const [messages, setMessages] = React.useState<Array<{ role: string; content: string; provider?: string }>>([]);
  const [loading, setLoading] = React.useState(false);
  const [aiManager, setAiManager] = React.useState<AIManager | null>(null);
  const [service, setService] = React.useState<any>(null);
  const [config, setConfig] = React.useState({
    geminiApiKey: '',
    ollamaUrl: 'http://localhost:11434',
    ollamaModel: 'llama3'
  });
  const [showConfig, setShowConfig] = React.useState(false);
  
  // Log para debugging
  React.useEffect(() => {
    console.log('🔧 showConfig cambió a:', showConfig);
  }, [showConfig]);
  const [currentProvider, setCurrentProvider] = React.useState('No configurado');

  React.useEffect(() => {
    // Cargar configuración guardada
    const savedConfig = localStorage.getItem('ai-config');
    if (savedConfig) {
      const parsed = JSON.parse(savedConfig);
      setConfig(parsed);
      initializeAI(parsed);
    }

    // Inicializar servicio según la aplicación
    switch (host) {
      case Office.HostType.Excel:
        setService(new ExcelService());
        break;
      case Office.HostType.Word:
        setService(new WordService());
        break;
      case Office.HostType.PowerPoint:
        setService(new PowerPointService());
        break;
    }
  }, [host]);

  const initializeAI = async (cfg: any) => {
    try {
      const manager = new AIManager(cfg);
      await manager.initialize();
      setAiManager(manager);
      setCurrentProvider(manager.getCurrentProvider());
      addMessage('system', `✅ Sistema iniciado. Usando: ${manager.getCurrentProvider()}`);
    } catch (error: any) {
      addMessage('system', `⚠️ ${error.message}`);
    }
  };

  const addMessage = (role: string, content: string, provider?: string) => {
    setMessages(prev => [...prev, { role, content, provider }]);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    console.log('📝 handleSubmit ejecutado');
    console.log('Prompt:', prompt);
    console.log('Loading:', loading);
    
    if (!prompt.trim() || loading) {
      console.log('⚠️ Submit cancelado - prompt vacío o cargando');
      return;
    }

    const userPrompt = prompt;
    setPrompt('');
    addMessage('user', userPrompt);
    setLoading(true);
    console.log('✅ Procesando mensaje...');

    try {
      if (!aiManager) {
        throw new Error('Sistema de IA no inicializado. Configura tu API key.');
      }

      // Obtener contexto de la aplicación actual
      let context = '';
      if (service) {
        try {
          context = await service.getContext();
        } catch {
          context = 'No se pudo obtener el contexto';
        }
      }

      // Construir prompt mejorado
      const enhancedPrompt = `
Eres un asistente para ${getHostName(host)}. 
El usuario quiere: "${userPrompt}"

Analiza la solicitud y responde en este formato JSON:
{
  "accion": "descripción de la acción a realizar",
  "comando": "comando específico para ejecutar",
  "explicacion": "breve explicación de lo que harás"
}

Contexto actual: ${context}
`;

      // Obtener respuesta de la IA
      const result = await aiManager.generateResponse(enhancedPrompt, context);
      
      // Mostrar respuesta de la IA
      addMessage('assistant', result.response, result.provider);
      
      if (result.fallbackUsed) {
        addMessage('system', '⚠️ Se agotó la cuota de Gemini, usando Ollama');
      }

      // Ejecutar comando si hay un servicio disponible
      if (service) {
        try {
          const commandResult = await service.executeCommand(userPrompt);
          addMessage('system', commandResult);
        } catch (error: any) {
          addMessage('system', `Error: ${error.message}`);
        }
      }

      setCurrentProvider(aiManager.getCurrentProvider());
    } catch (error: any) {
      addMessage('system', `❌ Error: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const saveConfig = () => {
    localStorage.setItem('ai-config', JSON.stringify(config));
    initializeAI(config);
    setShowConfig(false);
  };

  const getHostName = (hostType: Office.HostType): string => {
    switch (hostType) {
      case Office.HostType.Excel: return 'Excel';
      case Office.HostType.Word: return 'Word';
      case Office.HostType.PowerPoint: return 'PowerPoint';
      default: return 'Office';
    }
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>🤖 AI Copilot</h1>
        <p className="host-info">
          {getHostName(host)} | {currentProvider}
        </p>
        <button 
          className="config-btn" 
          onClick={() => {
            console.log('🖱️ Click en botón de configuración');
            console.log('showConfig actual:', showConfig);
            setShowConfig(!showConfig);
          }}
        >
          ⚙️
        </button>
      </header>

      {showConfig && (
        <div className="config-panel">
          <h3>Configuración</h3>
          <div className="config-field">
            <label>Gemini API Key:</label>
            <input
              type="password"
              value={config.geminiApiKey}
              onChange={(e) => setConfig({ ...config, geminiApiKey: e.target.value })}
              placeholder="AIza..."
            />
          </div>
          <div className="config-field">
            <label>Ollama URL:</label>
            <input
              type="text"
              value={config.ollamaUrl}
              onChange={(e) => setConfig({ ...config, ollamaUrl: e.target.value })}
              placeholder="http://localhost:11434"
            />
          </div>
          <div className="config-field">
            <label>Ollama Model:</label>
            <input
              type="text"
              value={config.ollamaModel}
              onChange={(e) => setConfig({ ...config, ollamaModel: e.target.value })}
              placeholder="llama3"
            />
          </div>
          <button className="save-btn" onClick={saveConfig}>
            💾 Guardar
          </button>
          <div className="config-help">
            <p>📚 <a href="https://ai.google.dev/" target="_blank">Obtener Gemini API Key</a></p>
            <p>📚 <a href="https://ollama.com/download" target="_blank">Descargar Ollama</a></p>
          </div>
        </div>
      )}

      <div className="messages-container">
        {messages.length === 0 && (
          <div className="welcome-message">
            <h2>¡Bienvenido! 👋</h2>
            <p>Escribe qué quieres hacer en {getHostName(host)}</p>
            <div className="examples">
              <h3>Ejemplos:</h3>
              {host === Office.HostType.Excel && (
                <>
                  <div className="example">"Crea una hoja llamada Ventas 2024"</div>
                  <div className="example">"Suma la columna A"</div>
                  <div className="example">"Crea un gráfico"</div>
                </>
              )}
              {host === Office.HostType.Word && (
                <>
                  <div className="example">"Escribe un párrafo sobre IA"</div>
                  <div className="example">"Crea una tabla de 3x3"</div>
                  <div className="example">"Agrega un título"</div>
                </>
              )}
              {host === Office.HostType.PowerPoint && (
                <>
                  <div className="example">"Crea un slide con título Resultados"</div>
                  <div className="example">"Agrega un rectángulo"</div>
                  <div className="example">"Crea 5 slides"</div>
                </>
              )}
            </div>
          </div>
        )}

        {messages.map((msg, idx) => (
          <div key={idx} className={`message message-${msg.role}`}>
            <div className="message-header">
              <span className="message-role">
                {msg.role === 'user' ? '👤 Tú' : msg.role === 'assistant' ? '🤖 AI' : '⚙️ Sistema'}
              </span>
              {msg.provider && <span className="provider-badge">{msg.provider}</span>}
            </div>
            <div className="message-content">{msg.content}</div>
          </div>
        ))}

        {loading && (
          <div className="message message-loading">
            <div className="loader">Procesando...</div>
          </div>
        )}
      </div>

      <form className="input-form" onSubmit={handleSubmit}>
        <input
          type="text"
          className="prompt-input"
          value={prompt}
          onChange={(e) => {
            console.log('⌨️ Escribiendo:', e.target.value);
            setPrompt(e.target.value);
          }}
          placeholder="¿Qué quieres hacer?"
          disabled={loading}
        />
        <button 
          type="submit" 
          className="send-btn" 
          disabled={loading || !prompt.trim()}
          onClick={() => console.log('🖱️ Click en botón enviar')}
        >
          ➤
        </button>
      </form>
    </div>
  );
};

export default App;

