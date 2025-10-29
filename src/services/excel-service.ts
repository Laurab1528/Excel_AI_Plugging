// Servicio para manejar operaciones de Excel
export class ExcelService {
  
  async executeCommand(command: string): Promise<string> {
    return Excel.run(async (context) => {
      try {
        const parsedCommand = this.parseCommand(command);
        
        switch (parsedCommand.action) {
          case 'crear_hoja':
            return await this.createSheet(context, parsedCommand.params);
          
          case 'agregar_datos':
            return await this.addData(context, parsedCommand.params);
          
          case 'insertar_formula':
            return await this.insertFormula(context, parsedCommand.params);
          
          case 'crear_grafico':
            return await this.createChart(context, parsedCommand.params);
          
          case 'formato':
            return await this.applyFormat(context, parsedCommand.params);
          
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
    
    // Detectar tipo de comando usando palabras clave
    if (lowerCommand.includes('crear') && (lowerCommand.includes('hoja') || lowerCommand.includes('sheet'))) {
      return { 
        action: 'crear_hoja', 
        params: { name: this.extractName(command) } 
      };
    }
    
    if (lowerCommand.includes('agregar') || lowerCommand.includes('insertar')) {
      return { 
        action: 'agregar_datos', 
        params: { data: command } 
      };
    }
    
    if (lowerCommand.includes('formula') || lowerCommand.includes('suma') || lowerCommand.includes('promedio')) {
      return { 
        action: 'insertar_formula', 
        params: { formula: command } 
      };
    }
    
    if (lowerCommand.includes('grafico') || lowerCommand.includes('chart')) {
      return { 
        action: 'crear_grafico', 
        params: { type: 'column' } 
      };
    }
    
    if (lowerCommand.includes('formato') || lowerCommand.includes('negrita') || lowerCommand.includes('color')) {
      return { 
        action: 'formato', 
        params: { command } 
      };
    }
    
    return { action: 'unknown', params: {} };
  }

  private extractName(text: string): string {
    // Intentar extraer nombre entre comillas
    const quotedMatch = text.match(/["']([^"']+)["']/);
    if (quotedMatch) return quotedMatch[1];
    
    // Buscar después de palabras clave
    const keywords = ['llamada', 'llamado', 'nombre', 'named'];
    for (const keyword of keywords) {
      const regex = new RegExp(`${keyword}\\s+["']?([\\w\\s]+)["']?`, 'i');
      const match = text.match(regex);
      if (match) return match[1].trim();
    }
    
    return 'Nueva Hoja';
  }

  private async createSheet(context: Excel.RequestContext, params: any): Promise<string> {
    const sheets = context.workbook.worksheets;
    const newSheet = sheets.add(params.name || 'Nueva Hoja');
    newSheet.activate();
    
    await context.sync();
    return `✅ Hoja "${params.name}" creada exitosamente`;
  }

  private async addData(context: Excel.RequestContext, params: any): Promise<string> {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Ejemplo: agregar datos en una fila
    const range = sheet.getRange('A1:C1');
    range.values = [['Nombre', 'Cantidad', 'Precio']];
    range.format.font.bold = true;
    
    await context.sync();
    return '✅ Datos agregados exitosamente';
  }

  private async insertFormula(context: Excel.RequestContext, params: any): Promise<string> {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange('A10');
    
    // Detectar tipo de fórmula
    if (params.formula.toLowerCase().includes('suma')) {
      range.formulas = [['=SUM(A1:A9)']];
    } else if (params.formula.toLowerCase().includes('promedio')) {
      range.formulas = [['=AVERAGE(A1:A9)']];
    }
    
    await context.sync();
    return '✅ Fórmula insertada exitosamente';
  }

  private async createChart(context: Excel.RequestContext, params: any): Promise<string> {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load('address');
    
    await context.sync();
    
    const chart = sheet.charts.add(
      Excel.ChartType.columnClustered,
      range,
      Excel.ChartSeriesBy.auto
    );
    
    chart.title.text = 'Gráfico Generado';
    
    await context.sync();
    return '✅ Gráfico creado exitosamente';
  }

  private async applyFormat(context: Excel.RequestContext, params: any): Promise<string> {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    
    if (params.command.toLowerCase().includes('negrita')) {
      range.format.font.bold = true;
    }
    
    if (params.command.toLowerCase().includes('color')) {
      range.format.font.color = 'blue';
    }
    
    await context.sync();
    return '✅ Formato aplicado exitosamente';
  }

  async getContext(): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      
      const range = sheet.getUsedRange();
      range.load('address,values');
      
      await context.sync();
      
      return `Hoja activa: ${sheet.name}\nRango usado: ${range.address}\nDatos: ${JSON.stringify(range.values)}`;
    });
  }
}

