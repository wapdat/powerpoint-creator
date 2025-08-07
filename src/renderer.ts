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
    primary: '#2C3E50',
    secondary: '#34495E',
    accent: '#3498DB',
    success: '#27AE60',
    warning: '#F39C12',
    danger: '#E74C3C',
    light: '#ECF0F1',
    dark: '#1A1A1A',
  },
  corporate: {
    primary: '#003366',
    secondary: '#005599',
    accent: '#0077CC',
    success: '#00AA44',
    warning: '#FFAA00',
    danger: '#CC0000',
    light: '#F5F5F5',
    dark: '#333333',
  },
  modern: {
    primary: '#6C63FF',
    secondary: '#4A47A3',
    accent: '#FF6584',
    success: '#00BFA6',
    warning: '#FFA726',
    danger: '#EF5350',
    light: '#FAFAFA',
    dark: '#212121',
  },
};

export class SlideRenderer {
  private pptx: PptxGenJS;
  private colorScheme: typeof COLOR_SCHEMES.professional;
  
  constructor(theme: string = 'professional') {
    this.pptx = new PptxGenJS();
    this.colorScheme = COLOR_SCHEMES[theme as keyof typeof COLOR_SCHEMES] || COLOR_SCHEMES.professional;
    this.setupPresentation();
  }
  
  /**
   * Setup default presentation properties
   */
  private setupPresentation(): void {
    this.pptx.author = 'pptx-auto-gen';
    this.pptx.company = 'Generated Presentation';
    this.pptx.revision = '1.0.0';
    this.pptx.subject = 'Professional Presentation';
    this.pptx.title = 'Presentation';
    
    // Set default layout
    this.pptx.layout = 'LAYOUT_16x9';
    
    // Define master slide
    this.pptx.defineSlideMaster({
      title: 'MASTER_SLIDE',
      background: { color: 'FFFFFF' },
      objects: [
        {
          placeholder: {
            options: {
              name: 'title',
              type: 'title',
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 1.5,
              fontSize: 32,
              bold: true,
              color: this.colorScheme.primary,
              fontFace: 'Arial',
            },
          },
        },
        {
          placeholder: {
            options: {
              name: 'body',
              type: 'body',
              x: 0.5,
              y: 2.5,
              w: 9,
              h: 4,
              fontSize: 18,
              color: this.colorScheme.dark,
              fontFace: 'Arial',
            },
          },
        },
        {
          line: {
            x: 0.5,
            y: 2.2,
            w: 9,
            h: 0,
            line: { color: this.colorScheme.accent, width: 2 },
          },
        },
      ],
    });
  }
  
  /**
   * Render a complete presentation
   */
  async renderPresentation(presentation: Presentation): Promise<Buffer> {
    // Set presentation metadata
    if (presentation.title) this.pptx.title = presentation.title;
    if (presentation.author) this.pptx.author = presentation.author;
    if (presentation.subject) this.pptx.subject = presentation.subject;
    if (presentation.company) this.pptx.company = presentation.company;
    
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
        throw new Error(`Unknown slide layout: ${slideData.layout}`);
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
    const slide = this.pptx.addSlide({ masterName: 'MASTER_SLIDE' });
    
    // Add title
    slide.addText(slideData.title, {
      placeholder: 'title',
    });
    
    // Process bullets with HTML support
    const bulletPoints = slideData.bullets.map((bullet, index) => {
      const level = slideData.level?.[index] || 0;
      return {
        text: this.processHtmlText(bullet),
        options: {
          bullet: { type: level === 0 ? 'bullet' : 'number' },
          indentLevel: level,
        },
      };
    });
    
    // Add bullets
    slide.addText(bulletPoints as any, {
      x: 0.5,
      y: 2.5,
      w: 9,
      h: 4,
      fontSize: 18,
      color: this.colorScheme.dark,
      fontFace: 'Arial',
      bullet: true,
      lineSpacing: 28,
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
    const slide = this.pptx.addSlide({ masterName: 'MASTER_SLIDE' });
    
    // Add title
    slide.addText(slideData.title, {
      placeholder: 'title',
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
    
    // Add image
    if (imageData) {
      const sizing = slideData.sizing || 'contain';
      const imageOptions: any = {
        x: 0.5,
        y: 2.5,
        w: 9,
        h: 3.5,
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
    
    // Add caption
    if (slideData.caption) {
      slide.addText(slideData.caption, {
        x: 0.5,
        y: 6.2,
        w: 9,
        h: 0.5,
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
    const slide = this.pptx.addSlide({ masterName: 'MASTER_SLIDE' });
    
    // Add title
    slide.addText(slideData.title, {
      placeholder: 'title',
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
    
    // Add chart
    slide.addChart(chartTypeMap[slideData.chartType], chartData, {
      x: 0.5,
      y: 2.5,
      w: 9,
      h: 4,
      showLegend: slideData.options?.showLegend !== false,
      legendPos: slideData.options?.legendPosition || 'b',
      showTitle: false,
      showValue: slideData.options?.showDataLabels,
      chartColors: slideData.options?.colors || [
        this.colorScheme.accent,
        this.colorScheme.success,
        this.colorScheme.warning,
        this.colorScheme.danger,
      ],
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
    const slide = this.pptx.addSlide({ masterName: 'MASTER_SLIDE' });
    
    // Add title
    slide.addText(slideData.title, {
      placeholder: 'title',
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
        },
      })));
    });
    
    // Add table
    slide.addTable(rows as any, {
      x: 0.5,
      y: 2.5,
      w: 9,
      h: 4,
      fontSize: slideData.styling?.fontSize || 14,
      fontFace: 'Arial',
      border: {
        type: 'solid',
        color: slideData.styling?.borderColor || this.colorScheme.secondary,
        pt: slideData.styling?.borderWidth || 1,
      },
      autoPage: false,
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
    const slide = this.pptx.addSlide({ masterName: 'MASTER_SLIDE' });
    
    // Add title
    slide.addText(slideData.title, {
      placeholder: 'title',
    });
    
    // Add content
    slide.addText(slideData.content, {
      x: 0.5,
      y: 2.5,
      w: 9,
      h: 4,
      fontSize: 16,
      color: this.colorScheme.dark,
      fontFace: 'Arial',
      valign: 'top',
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
  
  /**
   * Process HTML text for formatting
   */
  private processHtmlText(text: string): string {
    // Basic HTML tag processing
    return text
      .replace(/<strong>(.*?)<\/strong>/g, '**$1**')
      .replace(/<em>(.*?)<\/em>/g, '*$1*')
      .replace(/<u>(.*?)<\/u>/g, '_$1_')
      .replace(/<br\s*\/?>/g, '\n')
      .replace(/<[^>]*>/g, ''); // Remove any remaining HTML tags
  }
}