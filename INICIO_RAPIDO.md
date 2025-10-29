# 🚀 Guía de Inicio Rápido

## En 5 Minutos

### 1️⃣ Instalar dependencias
```bash
npm install
```

### 2️⃣ Obtener Gemini API Key (GRATIS)

1. Ve a [https://ai.google.dev/](https://ai.google.dev/)
2. Haz clic en "Get API Key"
3. Copia tu API key

### 3️⃣ Iniciar servidor
```bash
npm start
```

### 4️⃣ Cargar en Office

**Excel:**
1. Abre Excel
2. **Insertar** > **Complementos** > **Mis complementos** > **Cargar complemento personalizado**
3. Selecciona `manifests/manifest-excel.xml`

**Word:**
1. Abre Word
2. **Insertar** > **Complementos** > **Mis complementos** > **Cargar complemento personalizado**
3. Selecciona `manifests/manifest-word.xml`

**PowerPoint:**
1. Abre PowerPoint
2. **Insertar** > **Complementos** > **Mis complementos** > **Cargar complemento personalizado**
3. Selecciona `manifests/manifest-powerpoint.xml`

### 5️⃣ Configurar el Add-in

1. Haz clic en "Abrir AI Copilot" en la cinta
2. Haz clic en el ícono ⚙️
3. Pega tu Gemini API Key
4. Haz clic en **Guardar**

## ✅ ¡Listo!

Ahora puedes escribir comandos como:
- "Crea una hoja llamada Ventas 2024"
- "Escribe un párrafo sobre IA"
- "Crea un slide con título Resultados"

---

## 🆓 Sin Querer Usar Gemini?

### Instalar Ollama (100% Local y Gratis)

```bash
# Descargar e instalar desde: https://ollama.com/download

# Luego ejecutar:
ollama pull llama3

# Verificar que funciona:
ollama list
```

En la configuración del add-in:
- **Ollama URL**: `http://localhost:11434`
- **Ollama Model**: `llama3`

---

## ❓ Problemas Comunes

### "Sistema de IA no inicializado"
→ Verifica que hayas ingresado tu API key en la configuración

### El add-in no aparece
→ Asegúrate de que `npm start` esté corriendo
→ Reinicia Office

### Error de certificado SSL
→ Es normal en desarrollo, acepta el certificado en el navegador

---

📖 Para más detalles, consulta el [README.md](README.md) completo

