/**
 * Slide renderer module using PptxGenJS
 * Handles creation of presentations from scratch with professional styling
 */

import PptxGenJS from 'pptxgenjs';
import * as fs from 'fs';
import axios from 'axios';
import {
  Presentation,
  Slide,
  TitleSlide,
  TextSlide,
  ImageSlide,
  ChartSlide,
  TableSlide,
  NotesSlide,
  CustomSlide,
  ChartType,
  SlideElement,
} from './types';

/**
 * Professional color schemes for business presentations
 */
const COLOR_SCHEMES = {
  professional: {
    primary: '#2C3E50',      // Dark blue-gray
    secondary: '#546E7A',     // Slate gray
    accent: '#1976D2',        // Professional blue
    success: '#388E3C',       // Forest green
    warning: '#F57C00',       // Muted orange
    danger: '#D32F2F',        // Muted red
    light: '#F5F5F5',         // Light gray
    dark: '#263238',          // Dark gray
    // Chart colors - professional palette
    chart1: '#37474F',        // Charcoal
    chart2: '#607D8B',        // Blue-gray
    chart3: '#90A4AE',        // Light blue-gray
    chart4: '#455A64',        // Dark slate
    chart5: '#78909C',        // Medium blue-gray
    chart6: '#B0BEC5',        // Light slate
  },
  corporate: {
    primary: '#1E3A5F',       // Navy blue
    secondary: '#4A5F7A',     // Steel blue
    accent: '#2E5090',        // Corporate blue
    success: '#2E7D32',       // Professional green
    warning: '#ED6C02',       // Amber
    danger: '#C62828',        // Corporate red
    light: '#FAFAFA',         // Off white
    dark: '#1A1A1A',          // Near black
    // Chart colors - corporate palette
    chart1: '#1E3A5F',        // Navy
    chart2: '#4A5F7A',        // Steel blue
    chart3: '#7A8B99',        // Gray-blue
    chart4: '#2E5090',        // Royal blue
    chart5: '#5C7CAD',        // Soft blue
    chart6: '#9FAFC4',        // Light steel
  },
  modern: {
    primary: '#424242',       // Charcoal
    secondary: '#616161',     // Medium gray
    accent: '#1565C0',        // Modern blue
    success: '#2E7D32',       // Green
    warning: '#EF6C00',       // Orange
    danger: '#C62828',        // Red
    light: '#FAFAFA',         // White
    dark: '#212121',          // Dark
    // Chart colors - modern palette
    chart1: '#424242',        // Charcoal
    chart2: '#757575',        // Gray
    chart3: '#9E9E9E',        // Light gray
    chart4: '#1565C0',        // Blue
    chart5: '#42A5F5',        // Light blue
    chart6: '#BDBDBD',        // Silver
  },
};

export class SlideRenderer {
  private pptx: PptxGenJS;
  private colorScheme: any; // Dynamic color scheme
  
  constructor(theme: string = 'professional') {
    this.pptx = new PptxGenJS();
    this.colorScheme = COLOR_SCHEMES[theme as keyof typeof COLOR_SCHEMES] || COLOR_SCHEMES.professional;
    this.setupPresentation();
  }
  
  /**
   * Setup default presentation properties
   */
  private setupPresentation(): void {
    // Set minimal metadata to avoid potential XML issues
    this.pptx.author = 'PowerPoint Creator';
    this.pptx.title = 'Presentation';
    
    // Set default layout
    this.pptx.layout = 'LAYOUT_16x9';
  }
  
  /**
   * Render a complete presentation
   */
  async renderPresentation(presentation: Presentation): Promise<Buffer> {
    // Set only essential metadata to avoid XML issues
    if (presentation.title) this.pptx.title = presentation.title;
    if (presentation.author) this.pptx.author = presentation.author;
    
    // Apply theme if provided
    if (presentation.theme) {
      this.applyTheme(presentation.theme);
    }
    
    // Render each slide
    for (const slideData of presentation.slides) {
      await this.renderSlide(slideData);
    }
    
    // Generate and return buffer
    const buffer = await this.pptx.write({ outputType: 'nodebuffer' });
    return buffer as Buffer;
  }
  
  /**
   * Apply custom theme to presentation
   */
  private applyTheme(theme: any): void {
    if (theme.primaryColor) {
      this.colorScheme.primary = theme.primaryColor;
    }
    if (theme.secondaryColor) {
      this.colorScheme.secondary = theme.secondaryColor;
    }
    if (theme.accentColor) {
      this.colorScheme.accent = theme.accentColor;
    }
  }
  
  /**
   * Render individual slide based on layout type
   */
  private async renderSlide(slideData: Slide): Promise<void> {
    switch (slideData.layout) {
      case 'title':
        await this.renderTitleSlide(slideData as TitleSlide);
        break;
      case 'text':
        await this.renderTextSlide(slideData as TextSlide);
        break;
      case 'image':
        await this.renderImageSlide(slideData as ImageSlide);
        break;
      case 'chart':
        await this.renderChartSlide(slideData as ChartSlide);
        break;
      case 'table':
        await this.renderTableSlide(slideData as TableSlide);
        break;
      case 'notes':
        await this.renderNotesSlide(slideData as NotesSlide);
        break;
      case 'custom':
        await this.renderCustomSlide(slideData as CustomSlide);
        break;
      default:
        throw new Error(`Unknown slide layout: ${(slideData as any).layout}`);
    }
  }
  
  /**
   * Render title slide
   */
  private async renderTitleSlide(slideData: TitleSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Set background
    if (slideData.backgroundColor) {
      slide.background = { color: slideData.backgroundColor };
    } else {
      slide.background = { color: this.colorScheme.primary };
    }
    
    // Add title
    slide.addText(slideData.title, {
      x: 0.5,
      y: 2,
      w: 9,
      h: 1.5,
      fontSize: 44,
      bold: true,
      color: 'FFFFFF',
      align: 'center',
      fontFace: 'Arial',
    });
    
    // Add subtitle
    if (slideData.subtitle) {
      slide.addText(slideData.subtitle, {
        x: 0.5,
        y: 3.8,
        w: 9,
        h: 1,
        fontSize: 24,
        color: this.colorScheme.light,
        align: 'center',
        fontFace: 'Arial',
      });
    }
    
    // Add author and date
    if (slideData.author || slideData.date) {
      const footerText = [slideData.author, slideData.date].filter(Boolean).join(' | ');
      slide.addText(footerText, {
        x: 0.5,
        y: 5.5,
        w: 9,
        h: 0.5,
        fontSize: 14,
        color: this.colorScheme.light,
        align: 'center',
        fontFace: 'Arial',
      });
    }
    
    // Add speaker notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  
  /**
   * Render text slide with bullets
   */
  private async renderTextSlide(slideData: TextSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Add title with padding
    slide.addText(slideData.title, {
      x: 0.75,
      y: 0.4,
      w: 8.5,
      h: 0.8,
      fontSize: 22,
      bold: true,
      color: this.colorScheme.primary,
      fontFace: 'Arial',
    });
    
    // Process and clean bullet text
    const processedBullets: string[] = [];
    slideData.bullets.forEach((bullet) => {
      // Strip HTML tags and markdown for now to avoid XML issues
      let cleanText = bullet
        .replace(/<strong>(.*?)<\/strong>/gi, '$1')
        .replace(/<em>(.*?)<\/em>/gi, '$1')
        .replace(/<i>(.*?)<\/i>/gi, '$1')
        .replace(/<u>(.*?)<\/u>/gi, '$1')
        .replace(/\*\*(.*?)\*\*/g, '$1')
        .replace(/\*(.*?)\*/g, '$1')
        .replace(/__(.*?)__/g, '$1')
        .replace(/<[^>]*>/g, '');
      
      processedBullets.push(cleanText);
    });
    
    // Add bullets positioned properly
    const bulletText = processedBullets.map(text => ({ text, options: { bullet: true } }));
    slide.addText(bulletText, {
      x: 0.75,
      y: 1.3,
      w: 8.5,
      h: 4.5,
      fontSize: 16,
      color: this.colorScheme.dark,
      fontFace: 'Arial',
      lineSpacing: 28
    });
    
    // Add speaker notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  
  /**
   * Render image slide
   */
  private async renderImageSlide(slideData: ImageSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Add title with padding
    slide.addText(slideData.title, {
      x: 0.75,
      y: 0.4,
      w: 8.5,
      h: 0.8,
      fontSize: 22,
      bold: true,
      color: this.colorScheme.primary,
      fontFace: 'Arial',
    });
    
    // Prepare image data
    let imageData: string | { data: string };
    
    if (slideData.imageUrl) {
      // Download image from URL
      try {
        const response = await axios.get(slideData.imageUrl, {
          responseType: 'arraybuffer',
        });
        const base64 = Buffer.from(response.data).toString('base64');
        imageData = `data:${response.headers['content-type']};base64,${base64}`;
      } catch (error) {
        console.error(`Failed to download image: ${slideData.imageUrl}`);
        imageData = ''; // Fallback to empty
      }
    } else if (slideData.imagePath) {
      // Read local image
      if (fs.existsSync(slideData.imagePath)) {
        imageData = slideData.imagePath;
      } else {
        console.error(`Image file not found: ${slideData.imagePath}`);
        imageData = '';
      }
    } else {
      imageData = '';
    }
    
    // Add image positioned higher
    if (imageData) {
      const sizing = slideData.sizing || 'contain';
      const imageOptions: any = {
        x: 1,
        y: 1.3,
        w: 8,
        h: slideData.caption ? 3.8 : 4.2,
      };
      
      if (sizing === 'contain') {
        imageOptions.sizing = { type: 'contain' };
      } else if (sizing === 'cover') {
        imageOptions.sizing = { type: 'cover' };
      }
      
      slide.addImage({
        ...imageOptions,
        path: imageData,
      });
    }
    
    // Add caption with proper positioning
    if (slideData.caption) {
      slide.addText(slideData.caption, {
        x: 0.5,
        y: 6.1,
        w: 9,
        h: 0.4,
        fontSize: 14,
        italic: true,
        color: this.colorScheme.secondary,
        align: 'center',
        fontFace: 'Arial',
      });
    }
    
    // Add speaker notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  
  /**
   * Render chart slide
   */
  private async renderChartSlide(slideData: ChartSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Add title with padding
    slide.addText(slideData.title, {
      x: 0.75,
      y: 0.4,
      w: 8.5,
      h: 0.8,
      fontSize: 22,
      bold: true,
      color: this.colorScheme.primary,
      fontFace: 'Arial',
    });
    
    // Map chart type to PptxGenJS format
    const chartTypeMap: Record<ChartType, any> = {
      bar: this.pptx.ChartType.bar,
      line: this.pptx.ChartType.line,
      pie: this.pptx.ChartType.pie,
      area: this.pptx.ChartType.area,
      scatter: this.pptx.ChartType.scatter,
      doughnut: this.pptx.ChartType.doughnut,
      radar: this.pptx.ChartType.radar,
    };
    
    // Prepare chart data
    const chartData = slideData.data.datasets.map(dataset => ({
      name: dataset.label,
      labels: slideData.data.labels,
      values: dataset.data,
    }));
    
    // Add chart positioned higher to fit properly
    slide.addChart(chartTypeMap[slideData.chartType], chartData, {
      x: 1,
      y: 1.3,
      w: 8,
      h: 3.8,
      showLegend: slideData.options?.showLegend !== false,
      legendPos: (slideData.options?.legendPosition === 'top' ? 't' : 
                   slideData.options?.legendPosition === 'bottom' ? 'b' :
                   slideData.options?.legendPosition === 'left' ? 'l' :
                   slideData.options?.legendPosition === 'right' ? 'r' : 'r') as any,
      legendFontSize: 10,
      showTitle: false,
      showValue: slideData.options?.showDataLabels,
      chartColors: slideData.options?.colors || [
        this.colorScheme.chart1 || this.colorScheme.primary,
        this.colorScheme.chart2 || this.colorScheme.secondary,
        this.colorScheme.chart3 || this.colorScheme.accent,
        this.colorScheme.chart4 || this.colorScheme.success,
        this.colorScheme.chart5 || this.colorScheme.warning,
        this.colorScheme.chart6 || this.colorScheme.danger,
      ],
      valAxisLabelFontSize: 9,
      catAxisLabelFontSize: 9,
      dataLabelFontSize: 8,
    });
    
    // Add speaker notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  
  /**
   * Render table slide
   */
  private async renderTableSlide(slideData: TableSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Add title with padding
    slide.addText(slideData.title, {
      x: 0.75,
      y: 0.4,
      w: 8.5,
      h: 0.8,
      fontSize: 22,
      bold: true,
      color: this.colorScheme.primary,
      fontFace: 'Arial',
    });
    
    // Prepare table data
    const rows = [];
    
    // Add headers if provided
    if (slideData.headers) {
      rows.push(slideData.headers.map(header => ({
        text: header,
        options: {
          bold: true,
          color: slideData.styling?.headerTextColor || 'FFFFFF',
          fill: slideData.styling?.headerBackground || this.colorScheme.primary,
          align: 'center',
          valign: 'middle',
        },
      })));
    }
    
    // Add data rows
    slideData.tableData.forEach((row, rowIndex) => {
      const isAlternate = slideData.styling?.alternateRows && rowIndex % 2 === 1;
      rows.push(row.map(cell => ({
        text: String(cell),
        options: {
          color: this.colorScheme.dark,
          fill: isAlternate ? this.colorScheme.light : 'FFFFFF',
          align: 'left',
          valign: 'middle',
        },
      })));
    });
    
    // Calculate optimal row height
    const totalRows = rows.length;
    const availableHeight = 4.2;
    const rowHeight = Math.min(availableHeight / totalRows, 0.4);
    
    // Add table positioned higher
    slide.addTable(rows as any, {
      x: 0.75,
      y: 1.3,
      w: 8.5,
      h: 4.2,
      fontSize: slideData.styling?.fontSize || 14,
      fontFace: 'Arial',
      border: {
        type: 'solid',
        color: slideData.styling?.borderColor || this.colorScheme.secondary,
        pt: slideData.styling?.borderWidth || 1,
      },
      autoPage: false,
      rowH: rowHeight,
      margin: [0.1, 0.1, 0.1, 0.1],
    });
    
    // Add speaker notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  
  /**
   * Render notes slide
   */
  private async renderNotesSlide(slideData: NotesSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Add title with padding
    slide.addText(slideData.title, {
      x: 0.75,
      y: 0.4,
      w: 8.5,
      h: 0.8,
      fontSize: 22,
      bold: true,
      color: this.colorScheme.primary,
      fontFace: 'Arial',
    });
    
    // Add content positioned properly
    slide.addText(slideData.content, {
      x: 0.75,
      y: 1.3,
      w: 8.5,
      h: 4.5,
      fontSize: 16,
      color: this.colorScheme.dark,
      fontFace: 'Arial',
      valign: 'top',
      lineSpacing: 24,
    });
    
    // Add as speaker notes as well
    slide.addNotes(slideData.content);
  }
  
  /**
   * Render custom slide with multiple elements
   */
  private async renderCustomSlide(slideData: CustomSlide): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Add title if provided
    if (slideData.title) {
      slide.addText(slideData.title, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 1,
        fontSize: 32,
        bold: true,
        color: this.colorScheme.primary,
        fontFace: 'Arial',
      });
    }
    
    // Add each custom element
    for (const element of slideData.elements) {
      await this.renderCustomElement(slide, element);
    }
    
    // Add speaker notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  
  /**
   * Render individual custom element
   */
  private async renderCustomElement(slide: any, element: SlideElement): Promise<void> {
    const options = {
      x: element.x,
      y: element.y,
      w: element.width,
      h: element.height,
      ...element.styling,
    };
    
    switch (element.type) {
      case 'text':
        slide.addText(element.content, options);
        break;
        
      case 'image':
        slide.addImage({
          ...options,
          path: element.content,
        });
        break;
        
      case 'shape':
        slide.addShape(element.content.type || 'rect', {
          ...options,
          fill: element.content.fill || this.colorScheme.accent,
          line: element.content.line,
        });
        break;
        
      case 'chart':
        slide.addChart(element.content.type, element.content.data, options);
        break;
        
      case 'table':
        slide.addTable(element.content, options);
        break;
        
      default:
        console.warn(`Unknown element type: ${element.type}`);
    }
  }
  
  
  
}