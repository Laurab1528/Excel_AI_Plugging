# 🤖 AI Copilot para Office

Un complemento de Office potenciado por IA que funciona en **Excel**, **Word** y **PowerPoint**. Usa comandos en lenguaje natural para automatizar tareas y crear contenido.

## ✨ Características

- 🔄 **Multi-aplicación**: Funciona en Excel, Word y PowerPoint
- 🧠 **IA Dual**: Usa Gemini API (gratis) y fallback automático a Ollama
- 💬 **Lenguaje Natural**: Escribe comandos como hablarías normalmente
- 🚀 **Sin costos iniciales**: Usa APIs gratuitas
- 🔒 **Privacidad**: Opción de usar Ollama 100% local

## 🎯 Ejemplos de Uso

### Excel
```
"Crea una hoja llamada Ventas 2024"
"Suma la columna A y ponlo en A10"
"Crea un gráfico con los datos"
"Aplica formato de negrita a todo"
```

### Word
```
"Escribe un párrafo sobre inteligencia artificial"
"Crea una tabla de 3x4"
"Agrega un título llamado 'Informe Anual'"
"Crea una lista con 5 items"
```

### PowerPoint
```
"Crea un slide con título 'Resultados Q1'"
"Agrega un rectángulo en el slide actual"
"Crea 5 slides para una presentación"
```

## 📋 Requisitos Previos

- **Node.js** 16 o superior
- **npm** o **yarn**
- **Office 365** o **Office 2016+** (Windows/Mac)
- *(Opcional)* Cuenta Google para Gemini API (gratis)
- *(Opcional)* Ollama instalado para modo offline

## 🚀 Instalación

### 1. Clonar e instalar dependencias

```bash
# Clonar el proyecto (si es un repositorio)
git clone <tu-repo>
cd GPT_Excel

# Instalar dependencias
npm install
```

### 2. Configurar APIs de IA

#### Opción A: Gemini API (Recomendado - Gratis)

1. Ve a [Google AI Studio](https://ai.google.dev/)
2. Inicia sesión con tu cuenta Google
3. Haz clic en "Get API Key"
4. Copia tu API key

#### Opción B: Ollama (100% Local y Gratis)

1. Descarga Ollama: [https://ollama.com/download](https://ollama.com/download)
2. Instala Ollama
3. Abre terminal y ejecuta:
```bash
ollama pull llama3
```

> **Nota**: Puedes usar ambas opciones. El sistema intentará Gemini primero y cambiará a Ollama automáticamente si se agota la cuota.

### 3. Iniciar el servidor de desarrollo

```bash
npm start
```

Esto iniciará un servidor HTTPS en `https://localhost:3000`

### 4. Cargar el Add-in en Office

#### Excel

1. Abre Excel
2. Ve a **Insertar** > **Complementos** > **Mis complementos**
3. Selecciona **Complementos compartidos**
4. Haz clic en **Cargar complemento personalizado**
5. Navega a `manifests/manifest-excel.xml` y selecciónalo
6. Haz clic en **Cargar**

#### Word

1. Abre Word
2. Ve a **Insertar** > **Complementos** > **Mis complementos**
3. Selecciona **Complementos compartidos**
4. Haz clic en **Cargar complemento personalizado**
5. Navega a `manifests/manifest-word.xml` y selecciónalo
6. Haz clic en **Cargar**

#### PowerPoint

1. Abre PowerPoint
2. Ve a **Insertar** > **Complementos** > **Mis complementos**
3. Selecciona **Complementos compartidos**
4. Haz clic en **Cargar complemento personalizado**
5. Navega a `manifests/manifest-powerpoint.xml` y selecciónalo
6. Haz clic en **Cargar**

> **Nota para Mac**: El proceso es similar pero puede variar ligeramente según la versión de Office.

## ⚙️ Configuración

Una vez cargado el add-in:

1. Haz clic en el botón **"Abrir AI Copilot"** en la cinta
2. Haz clic en el icono de engranaje ⚙️ 
3. Configura tus credenciales:
   - **Gemini API Key**: Pega tu API key de Google
   - **Ollama URL**: Por defecto `http://localhost:11434`
   - **Ollama Model**: Modelo a usar (llama3, mistral, etc.)
4. Haz clic en **Guardar**

## 🎮 Uso

1. Abre el panel de AI Copilot desde la cinta de Office
2. Escribe tu comando en lenguaje natural
3. Presiona Enter o haz clic en ➤
4. El sistema procesará tu solicitud y ejecutará la acción

### Ejemplos Avanzados

#### Excel
```
"Crea una hoja de gastos con columnas: Fecha, Concepto, Monto"
"Calcula el promedio de la columna B"
"Crea un gráfico de barras con los datos de A1:B10"
```

#### Word
```
"Escribe una introducción de 200 palabras sobre cambio climático"
"Crea una tabla comparativa con 3 productos"
"Agrega un título de nivel 1 'Conclusiones'"
```

#### PowerPoint
```
"Crea 3 slides: Introducción, Desarrollo, Conclusión"
"Agrega un cuadro de texto con 'Objetivos del proyecto'"
```

## 🔧 Solución de Problemas

### El add-in no aparece en Office

- Asegúrate de que el servidor esté corriendo (`npm start`)
- Verifica que el manifiesto esté correctamente cargado
- Reinicia la aplicación de Office

### Error de certificado SSL

En desarrollo, es normal ver advertencias de certificado. Para Office:
- Windows: Acepta el certificado autofirmado
- Mac: Agrega el certificado a las "Utilidades de cadena de llaves"

### "Sistema de IA no inicializado"

- Verifica que hayas configurado al menos una API (Gemini o Ollama)
- Para Gemini: Asegúrate de que la API key es válida
- Para Ollama: Verifica que Ollama esté ejecutándose (`ollama list`)

### "Cuota de Gemini agotada"

- El sistema cambiará automáticamente a Ollama si está disponible
- O espera 24 horas para que se renueve tu cuota gratuita de Gemini
- O instala Ollama como alternativa permanente

### Ollama no responde

```bash
# Verifica que Ollama esté ejecutándose
ollama list

# Si no está, inícialo
ollama serve

# Verifica que el modelo esté descargado
ollama pull llama3
```

## 📦 Construir para Producción

```bash
npm run build
```

Esto generará los archivos optimizados en la carpeta `dist/`

## 🏗️ Estructura del Proyecto

```
GPT_Excel/
├── manifests/              # Manifiestos XML para Office
│   ├── manifest-excel.xml
│   ├── manifest-word.xml
│   └── manifest-powerpoint.xml
├── src/
│   ├── ai/                 # Sistema de IA
│   │   ├── types.ts
│   │   ├── gemini-provider.ts
│   │   ├── ollama-provider.ts
│   │   └── ai-manager.ts
│   ├── services/           # Servicios para cada app
│   │   ├── excel-service.ts
│   │   ├── word-service.ts
│   │   └── powerpoint-service.ts
│   └── taskpane/           # Interfaz de usuario
│       ├── taskpane.html
│       ├── taskpane.tsx
│       ├── App.tsx
│       └── App.css
├── assets/                 # Iconos
├── package.json
├── tsconfig.json
└── webpack.config.js
```

## 🤝 Contribuir

Las contribuciones son bienvenidas. Para cambios importantes:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📝 Notas Importantes

### Límites de API Gratuitas

- **Gemini**: 15 requests/minuto, 1,500/día (más que suficiente para uso personal)
- **Ollama**: Ilimitado (local)

### Privacidad

- **Gemini**: Los prompts se envían a Google (revisa su política de privacidad)
- **Ollama**: Todo se ejecuta localmente, cero envío de datos a internet

### Rendimiento

- **Gemini**: Respuestas rápidas (1-3 segundos)
- **Ollama**: Depende de tu hardware (2-10 segundos)

## 🆘 Soporte

¿Problemas o preguntas?

1. Revisa la sección de **Solución de Problemas** arriba
2. Busca en los [Issues del proyecto](si tienes repo)
3. Crea un nuevo Issue con detalles

## 📄 Licencia

MIT License - siéntete libre de usar este proyecto como quieras.

## 🎓 Recursos

- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Gemini API Documentation](https://ai.google.dev/docs)
- [Ollama Documentation](https://ollama.com/docs)
- [Office.js API Reference](https://learn.microsoft.com/en-us/javascript/api/overview)

---

**¡Hecho con ❤️ y mucha IA!**

