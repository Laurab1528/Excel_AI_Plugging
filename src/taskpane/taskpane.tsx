import * as React from 'react';
import * as ReactDOM from 'react-dom';
import App from './App';

/* global window */

console.log('✅ taskpane.js cargado');

let attemptCount = 0;
const maxAttempts = 30; // 3 segundos máximo

function mountApp(host: any) {
  const container = document.getElementById('root');
  const loadingIndicator = document.getElementById('loading-indicator');
  
  if (!container) {
    console.error('❌ No se encontró el contenedor #root');
    return;
  }
  
  console.log('✅ Montando React...', host);
  
  // Ocultar indicador de carga
  if (loadingIndicator) {
    loadingIndicator.style.display = 'none';
  }
  
  // Montar la aplicación React
  ReactDOM.render(
    <React.StrictMode>
      <App host={host} />
    </React.StrictMode>,
    container
  );
  
  console.log('✅ Aplicación montada correctamente');
}

function initializeApp() {
  attemptCount++;
  
  // Si Office.js está disponible, usarlo normalmente
  if (typeof (window as any).Office !== 'undefined') {
    console.log('✅ Office.js detectado');
    (window as any).Office.onReady((info: any) => {
      console.log('✅ Office.onReady ejecutado', info);
      mountApp(info.host);
    });
    return;
  }
  
  // Si llegamos al límite, montar de todos modos sin Office.js
  if (attemptCount >= maxAttempts) {
    console.warn('⚠️ Office.js no disponible - Montando app en modo limitado');
    mountApp('Excel' as any); // Valor por defecto
    return;
  }
  
  // Seguir esperando
  console.log(`⏳ Esperando Office.js... (${attemptCount}/${maxAttempts})`);
  setTimeout(initializeApp, 100);
}

// Iniciar cuando el DOM esté listo
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeApp);
} else {
  initializeApp();
}

