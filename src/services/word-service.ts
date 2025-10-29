// Servicio para manejar operaciones de Word
export class WordService {
  
  async executeCommand(command: string): Promise<string> {
    return Word.run(async (context) => {
      try {
        const parsedCommand = this.parseCommand(command);
        
        switch (parsedCommand.action) {
          case 'insertar_texto':
            return await this.insertText(context, parsedCommand.params);
          
          case 'crear_tabla':
            return await this.createTable(context, parsedCommand.params);
          
          case 'aplicar_estilo':
            return await this.applyStyle(context, parsedCommand.params);
          
          case 'insertar_imagen':
            return await this.insertImage(context, parsedCommand.params);
          
          case 'crear_lista':
            return await this.createList(context, parsedCommand.params);
          
          default:
            return `Comando no reconocido: ${parsedCommand.action}`;
        }
      } catch (error: any) {
        return `Error ejecutando comando: ${error.message}`;
      }
    });
  }

  private parseCommand(command: string): { action: string; params: any } {
    const lowerCommand = command.toLowerCase();
    
    if (lowerCommand.includes('escribir') || lowerCommand.includes('texto') || lowerCommand.includes('párrafo')) {
      return { 
        action: 'insertar_texto', 
        params: { text: this.extractContent(command) } 
      };
    }
    
    if (lowerCommand.includes('tabla')) {
      return { 
        action: 'crear_tabla', 
        params: this.extractTableParams(command) 
      };
    }
    
    if (lowerCommand.includes('título') || lowerCommand.includes('heading') || lowerCommand.includes('estilo')) {
      return { 
        action: 'aplicar_estilo', 
        params: { style: 'Heading1' } 
      };
    }
    
    if (lowerCommand.includes('imagen')) {
      return { 
        action: 'insertar_imagen', 
        params: {} 
      };
    }
    
    if (lowerCommand.includes('lista')) {
      return { 
        action: 'crear_lista', 
        params: { items: this.extractListItems(command) } 
      };
    }
    
    return { action: 'insertar_texto', params: { text: command } };
  }

  private extractContent(text: string): string {
    const quotedMatch = text.match(/["']([^"']+)["']/);
    return quotedMatch ? quotedMatch[1] : text;
  }

  private extractTableParams(text: string): any {
    const rowMatch = text.match(/(\d+)\s*(fila|row)/i);
    const colMatch = text.match(/(\d+)\s*(columna|column)/i);
    
    return {
      rows: rowMatch ? parseInt(rowMatch[1]) : 3,
      cols: colMatch ? parseInt(colMatch[1]) : 3
    };
  }

  private extractListItems(text: string): string[] {
    // Buscar items separados por comas o guiones
    const items = text.split(/[,\n-]/).map(item => item.trim()).filter(item => item.length > 0);
    return items.length > 0 ? items : ['Item 1', 'Item 2', 'Item 3'];
  }

  private async insertText(context: Word.RequestContext, params: any): Promise<string> {
    const body = context.document.body;
    body.insertParagraph(params.text || 'Texto insertado', Word.InsertLocation.end);
    
    await context.sync();
    return `✅ Texto insertado exitosamente`;
  }

  private async createTable(context: Word.RequestContext, params: any): Promise<string> {
    const body = context.document.body;
    const table = body.insertTable(
      params.rows || 3, 
      params.cols || 3, 
      Word.InsertLocation.end,
      [['Columna 1', 'Columna 2', 'Columna 3']]
    );
    
    table.styleBuiltIn = 'GridTable4';
    
    await context.sync();
    return `✅ Tabla de ${params.rows}x${params.cols} creada exitosamente`;
  }

  private async applyStyle(context: Word.RequestContext, params: any): Promise<string> {
    const body = context.document.body;
    body.insertParagraph('Título del Documento', Word.InsertLocation.start).styleBuiltIn = 'Heading1';
    
    await context.sync();
    return `✅ Estilo aplicado exitosamente`;
  }

  private async insertImage(context: Word.RequestContext, params: any): Promise<string> {
    return `⚠️ Inserción de imágenes requiere una URL. Proporciona una URL de imagen válida.`;
  }

  private async createList(context: Word.RequestContext, params: any): Promise<string> {
    const body = context.document.body;
    
    params.items.forEach((item: string, index: number) => {
      const paragraph = body.insertParagraph(item, Word.InsertLocation.end);
      if (index === 0) {
        paragraph.startNewList();
      }
    });
    
    await context.sync();
    return `✅ Lista con ${params.items.length} items creada exitosamente`;
  }

  async getContext(): Promise<string> {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      
      await context.sync();
      
      const preview = body.text.substring(0, 200);
      return `Documento actual (primeros 200 caracteres):\n${preview}...`;
    });
  }
}

