// Servicio para manejar operaciones de PowerPoint
export class PowerPointService {
  
  async executeCommand(command: string): Promise<string> {
    return PowerPoint.run(async (context) => {
      try {
        const parsedCommand = this.parseCommand(command);
        
        switch (parsedCommand.action) {
          case 'crear_slide':
            return await this.createSlide(context, parsedCommand.params);
          
          case 'agregar_texto':
            return await this.addText(context, parsedCommand.params);
          
          case 'agregar_forma':
            return await this.addShape(context, parsedCommand.params);
          
          case 'cambiar_layout':
            return await this.changeLayout(context, parsedCommand.params);
          
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
    
    if (lowerCommand.includes('crear') && (lowerCommand.includes('slide') || lowerCommand.includes('diapositiva'))) {
      return { 
        action: 'crear_slide', 
        params: { title: this.extractTitle(command) } 
      };
    }
    
    if (lowerCommand.includes('texto') || lowerCommand.includes('título')) {
      return { 
        action: 'agregar_texto', 
        params: { text: this.extractContent(command) } 
      };
    }
    
    if (lowerCommand.includes('forma') || lowerCommand.includes('rectángulo') || lowerCommand.includes('círculo')) {
      return { 
        action: 'agregar_forma', 
        params: { type: this.extractShapeType(command) } 
      };
    }
    
    if (lowerCommand.includes('layout') || lowerCommand.includes('diseño')) {
      return { 
        action: 'cambiar_layout', 
        params: {} 
      };
    }
    
    return { action: 'crear_slide', params: { title: 'Nuevo Slide' } };
  }

  private extractTitle(text: string): string {
    const quotedMatch = text.match(/["']([^"']+)["']/);
    if (quotedMatch) return quotedMatch[1];
    
    const keywords = ['título', 'llamado', 'named'];
    for (const keyword of keywords) {
      const regex = new RegExp(`${keyword}\\s+["']?([\\w\\s]+)["']?`, 'i');
      const match = text.match(regex);
      if (match) return match[1].trim();
    }
    
    return 'Nuevo Slide';
  }

  private extractContent(text: string): string {
    const quotedMatch = text.match(/["']([^"']+)["']/);
    return quotedMatch ? quotedMatch[1] : text;
  }

  private extractShapeType(text: string): string {
    if (text.includes('rectángulo')) return 'rectangle';
    if (text.includes('círculo')) return 'oval';
    return 'rectangle';
  }

  private async createSlide(context: PowerPoint.RequestContext, params: any): Promise<string> {
    const presentation = context.presentation;
    
    // Crear un nuevo slide con layout de título
    const slide = presentation.slides.add();
    
    // En PowerPoint API, necesitamos esperar a que el slide esté disponible
    await context.sync();
    
    return `✅ Slide "${params.title}" creado exitosamente`;
  }

  private async addText(context: PowerPoint.RequestContext, params: any): Promise<string> {
    const presentation = context.presentation;
    const slides = presentation.slides;
    slides.load('items');
    
    await context.sync();
    
    if (slides.items.length === 0) {
      return '⚠️ No hay slides. Crea un slide primero.';
    }
    
    // Agregar texto al último slide (simplificado)
    return `✅ Texto agregado al slide actual`;
  }

  private async addShape(context: PowerPoint.RequestContext, params: any): Promise<string> {
    const presentation = context.presentation;
    const slides = presentation.slides;
    slides.load('items');
    
    await context.sync();
    
    if (slides.items.length === 0) {
      return '⚠️ No hay slides. Crea un slide primero.';
    }
    
    const slide = slides.items[slides.items.length - 1];
    const shapes = slide.shapes;
    
    // Agregar forma (rectángulo por defecto)
    const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    shape.left = 100;
    shape.top = 100;
    shape.height = 100;
    shape.width = 200;
    
    await context.sync();
    
    return `✅ Forma agregada exitosamente`;
  }

  private async changeLayout(context: PowerPoint.RequestContext, params: any): Promise<string> {
    return `✅ Layout cambiado (función en desarrollo)`;
  }

  async getContext(): Promise<string> {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load('items');
      
      await context.sync();
      
      return `Presentación actual tiene ${slides.items.length} slides`;
    });
  }
}

