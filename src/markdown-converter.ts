/**
 * Markdown to PowerPoint Converter
 * Converts markdown documents to PowerPoint presentation JSON structure
 */

import { marked } from 'marked';
import matter from 'gray-matter';
import {
  Presentation,
  Slide,
  TitleSlide,
  TextSlide,
  ChartSlide,
  TableSlide,
  ImageSlide,
  NotesSlide,
  CustomSlide,
} from './types';

/**
 * Options for markdown to PowerPoint conversion
 */
export interface MarkdownConverterOptions {
  /** Maximum bullet points per slide */
  slidesPerPage?: number;
  /** Enable automatic content splitting */
  autoSplit?: boolean;
  /** Theme name or custom theme object */
  theme?: string | any;
  /** Generate table of contents slide */
  includeTableOfContents?: boolean;
  /** Add slide numbers */
  slideNumbers?: boolean;
  /** Company logo URL or path */
  companyLogo?: string;
  /** Footer text for all slides */
  footerText?: string;
  /** Maximum characters per slide */
  maxCharsPerSlide?: number;
}

/**
 * Internal structure for tracking markdown parsing state
 */
interface ParsingContext {
  slides: Slide[];
  currentSlide?: Partial<Slide>;
  currentBullets: string[];
  currentLevel: number[];
  inCodeBlock: boolean;
  inTable: boolean;
  tableHeaders: string[];
  tableRows: string[][];
  metadata: any;
  sectionDepth: number;
  slideCount: number;
}

/**
 * Chart data detection patterns
 */
const CHART_PATTERNS = {
  CSV: /^[^,]+,[^,]+(,[^,]+)*$/m,
  JSON: /^\s*\{[\s\S]*\}\s*$/,
  TABLE_NUMERIC: /^\s*\|.*\|.*\|.*\|\s*$/m,
};

/**
 * Markdown to PowerPoint Converter Class
 */
export class MarkdownConverter {
  private options: MarkdownConverterOptions;
  private defaultOptions: MarkdownConverterOptions = {
    slidesPerPage: 6,
    autoSplit: true,
    theme: 'professional',
    includeTableOfContents: false,
    slideNumbers: false,
    maxCharsPerSlide: 500,
  };

  constructor(options: MarkdownConverterOptions = {}) {
    this.options = { ...this.defaultOptions, ...options };
  }

  /**
   * Convert markdown content to presentation JSON
   * @param markdownContent - The markdown content to convert
   * @returns Presentation JSON structure
   */
  async convert(markdownContent: string): Promise<Presentation> {
    // Parse frontmatter
    const { data: frontmatter, content } = matter(markdownContent);
    
    // Initialize parsing context
    const context: ParsingContext = {
      slides: [],
      currentBullets: [],
      currentLevel: [],
      inCodeBlock: false,
      inTable: false,
      tableHeaders: [],
      tableRows: [],
      metadata: frontmatter,
      sectionDepth: 0,
      slideCount: 0,
    };

    // Set up custom renderer
    const renderer = this.createCustomRenderer(context);
    marked.use({ renderer });

    // Parse markdown
    marked.parse(content);

    // Flush any remaining content
    this.flushCurrentSlide(context);

    // Build presentation structure
    const presentation: Presentation = {
      title: frontmatter.title || 'Presentation',
      author: frontmatter.author,
      company: frontmatter.company,
      subject: frontmatter.subject,
      slides: context.slides,
    };

    // Apply theme if specified
    if (frontmatter.theme || this.options.theme) {
      presentation.theme = this.parseTheme(frontmatter.theme || this.options.theme);
    }

    // Add table of contents if requested
    if (this.options.includeTableOfContents) {
      this.insertTableOfContents(presentation);
    }

    // Add slide numbers if requested
    if (this.options.slideNumbers) {
      this.addSlideNumbers(presentation);
    }

    return presentation;
  }

  /**
   * Create custom marked renderer for PowerPoint conversion
   */
  private createCustomRenderer(context: ParsingContext): any {
    const renderer = new marked.Renderer();

    // Handle headings
    renderer.heading = (text: string, level: number): string => {
      // Flush any pending content
      this.flushCurrentSlide(context);

      if (level === 1) {
        // H1 -> Title slide
        const titleSlide: TitleSlide = {
          layout: 'title',
          title: this.cleanText(text),
          subtitle: context.metadata.subtitle,
          author: context.metadata.author,
          date: context.metadata.date || new Date().toLocaleDateString(),
        };
        context.slides.push(titleSlide);
      } else if (level === 2) {
        // H2 -> Section divider
        const sectionSlide: TitleSlide = {
          layout: 'title',
          title: this.cleanText(text),
          backgroundColor: '#2C3E50',
        };
        context.slides.push(sectionSlide);
        context.sectionDepth = 2;
      } else if (level === 3) {
        // H3 -> Add as emphasized content or start new slide
        if (context.currentBullets.length > 0) {
          // Add as emphasized bullet
          context.currentBullets.push(`**${this.cleanText(text)}**`);
          context.currentLevel.push(0);
        } else {
          // Start new content slide
          context.currentSlide = {
            layout: 'text',
            title: this.cleanText(text),
          };
          context.currentBullets = [];
          context.currentLevel = [];
        }
      }

      return '';
    };

    // Handle paragraphs
    renderer.paragraph = (text: string): string => {
      // Check for special directives
      if (this.isDirective(text)) {
        this.processDirective(text, context);
        return '';
      }

      // Check for chart/data blocks
      if (this.isChartData(text)) {
        this.processChartData(text, context);
        return '';
      }

      // Add as bullet point
      if (context.currentSlide || context.currentBullets.length > 0) {
        const cleanedText = this.cleanText(text);
        if (cleanedText) {
          context.currentBullets.push(cleanedText);
          context.currentLevel.push(0);
          
          // Auto-split if too many bullets
          if (this.options.autoSplit && context.currentBullets.length >= this.options.slidesPerPage!) {
            this.flushCurrentSlide(context);
          }
        }
      }

      return '';
    };

    // Handle lists
    renderer.list = (_body: string, _ordered: boolean): string => {
      // Lists are handled by listitem
      return '';
    };

    renderer.listitem = (text: string): string => {
      if (!context.currentSlide) {
        context.currentSlide = {
          layout: 'text',
          title: 'Content',
        };
      }

      const cleanedText = this.cleanText(text);
      const level = this.detectIndentLevel(text);
      
      context.currentBullets.push(cleanedText);
      context.currentLevel.push(level);

      // Auto-split if needed
      if (this.options.autoSplit && context.currentBullets.length >= this.options.slidesPerPage!) {
        this.flushCurrentSlide(context);
      }

      return '';
    };

    // Handle tables
    renderer.table = (header: string, body: string): string => {
      this.flushCurrentSlide(context);
      
      // Parse HTML table structure
      const headers = this.parseHTMLTableRow(header, 'th');
      
      // Parse all rows from body - body contains multiple <tr> elements
      const rows: string[][] = [];
      const rowMatches = body.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) || [];
      for (const rowHtml of rowMatches) {
        const rowData = this.parseHTMLTableRow(rowHtml, 'td');
        if (rowData.length > 0) {
          rows.push(rowData);
        }
      }

      // Check if table contains chart data
      if (this.isChartTable(headers, rows)) {
        const chartSlide = this.createChartFromTable(headers, rows, context);
        context.slides.push(chartSlide);
      } else {
        // Create table slide
        const tableSlide: TableSlide = {
          layout: 'table',
          title: context.currentSlide?.title || 'Data Table',
          headers: headers,
          tableData: rows,
          styling: {
            headerBackground: '#2C3E50',
            headerTextColor: '#FFFFFF',
            alternateRows: true,
          },
        };
        context.slides.push(tableSlide);
      }

      context.currentSlide = undefined;
      return '';
    };

    // Handle images
    renderer.image = (href: string, title: string | null, text: string): string => {
      this.flushCurrentSlide(context);

      const imageSlide: ImageSlide = {
        layout: 'image',
        title: title || text || 'Image',
        imageUrl: href.startsWith('http') ? href : undefined,
        imagePath: !href.startsWith('http') ? href : undefined,
        caption: text,
        sizing: 'contain',
      };

      context.slides.push(imageSlide);
      return '';
    };

    // Handle code blocks
    renderer.code = (code: string, language?: string): string => {
      // Check if it's chart data
      if (this.isChartData(code) || language === 'chart' || language === 'csv') {
        this.processChartData(code, context);
      } else {
        // Add as notes slide with monospace formatting
        this.flushCurrentSlide(context);
        
        const notesSlide: NotesSlide = {
          layout: 'notes',
          title: `Code: ${language || 'snippet'}`,
          content: code,
        };
        
        context.slides.push(notesSlide);
      }
      return '';
    };

    // Handle blockquotes
    renderer.blockquote = (quote: string): string => {
      this.flushCurrentSlide(context);

      const notesSlide: NotesSlide = {
        layout: 'notes',
        title: 'Note',
        content: this.cleanText(quote),
      };

      context.slides.push(notesSlide);
      return '';
    };

    // Handle horizontal rules (slide breaks)
    renderer.hr = (): string => {
      this.flushCurrentSlide(context);
      return '';
    };

    // Handle HTML comments for directives
    renderer.html = (html: string): string => {
      if (html.includes('slide:') || html.includes('notes:')) {
        this.processDirective(html, context);
      }
      return '';
    };

    return renderer;
  }

  /**
   * Flush current slide to slides array
   */
  private flushCurrentSlide(context: ParsingContext): void {
    if (context.currentBullets.length > 0) {
      const textSlide: TextSlide = {
        layout: 'text',
        title: context.currentSlide?.title || 'Content',
        bullets: context.currentBullets,
        level: context.currentLevel.length > 0 ? context.currentLevel : undefined,
      };

      // Add notes if present
      if (context.currentSlide && 'notes' in context.currentSlide) {
        textSlide.notes = (context.currentSlide as any).notes;
      }

      context.slides.push(textSlide);
      context.currentBullets = [];
      context.currentLevel = [];
      context.currentSlide = undefined;
    } else if (context.currentSlide && context.currentSlide.layout) {
      // Don't add empty text slides
      if (context.currentSlide.layout === 'text' && (!('bullets' in context.currentSlide) || (context.currentSlide as TextSlide).bullets?.length === 0)) {
        // Skip empty text slides
        context.currentSlide = undefined;
        return;
      }
      context.slides.push(context.currentSlide as Slide);
      context.currentSlide = undefined;
    }
  }

  /**
   * Clean text by removing HTML tags and extra whitespace
   */
  private cleanText(text: string): string {
    return text
      .replace(/<[^>]*>/g, '') // Remove HTML tags
      .replace(/\n+/g, ' ') // Replace newlines with spaces
      .replace(/\s+/g, ' ') // Normalize whitespace
      .trim();
  }

  /**
   * Check if text is a directive
   */
  private isDirective(text: string): boolean {
    return /<!--\s*(slide:|notes:|chart:)/.test(text) || /^:::(slide|notes|chart)/.test(text);
  }

  /**
   * Process directive comments
   */
  private processDirective(text: string, context: ParsingContext): void {
    // Extract directive type and content
    const slideMatch = text.match(/<!--\s*slide:\s*(.+?)\s*-->/);
    const notesMatch = text.match(/<!--\s*notes:\s*(.+?)\s*-->/);
    const chartMatch = text.match(/<!--\s*chart:\s*(.+?)\s*-->/);

    if (slideMatch) {
      // Parse slide directive
      try {
        const params = this.parseDirectiveParams(slideMatch[1]);
        if (params.type === 'custom') {
          this.flushCurrentSlide(context);
          const customSlide: CustomSlide = {
            layout: 'custom',
            title: params.title || 'Custom Slide',
            elements: params.elements || [],
          };
          context.slides.push(customSlide);
        }
      } catch (e) {
        console.warn('Failed to parse slide directive:', e);
      }
    } else if (notesMatch) {
      // Add speaker notes to current or next slide
      const notes = notesMatch[1].trim();
      if (context.currentSlide) {
        (context.currentSlide as any).notes = notes;
      } else if (context.slides.length > 0) {
        (context.slides[context.slides.length - 1] as any).notes = notes;
      }
    } else if (chartMatch) {
      // Process chart directive
      const params = this.parseDirectiveParams(chartMatch[1]);
      context.currentSlide = {
        layout: 'chart',
        chartType: params.type || 'bar',
      };
    }
  }

  /**
   * Parse directive parameters
   */
  private parseDirectiveParams(params: string): any {
    const result: any = {};
    const regex = /(\w+)="([^"]+)"/g;
    let match;

    while ((match = regex.exec(params)) !== null) {
      result[match[1]] = match[2];
    }

    return result;
  }

  /**
   * Check if text contains chart data
   */
  private isChartData(text: string): boolean {
    return CHART_PATTERNS.CSV.test(text) || 
           CHART_PATTERNS.JSON.test(text) ||
           CHART_PATTERNS.TABLE_NUMERIC.test(text);
  }

  /**
   * Process chart data from text
   */
  private processChartData(text: string, context: ParsingContext): void {
    this.flushCurrentSlide(context);

    try {
      let chartData: any;
      let chartType = 'bar';

      // Try to parse as JSON
      if (CHART_PATTERNS.JSON.test(text)) {
        chartData = JSON.parse(text);
      } else if (CHART_PATTERNS.CSV.test(text)) {
        // Parse CSV
        chartData = this.parseCSVToChartData(text);
      }

      if (chartData && chartData.datasets && chartData.datasets.length > 0) {
        const chartSlide: ChartSlide = {
          layout: 'chart',
          title: context.currentSlide?.title || 'Chart',
          chartType: chartData.type || chartType,
          data: chartData.data || chartData,
        };

        context.slides.push(chartSlide);
      }
    } catch (e) {
      console.warn('Failed to parse chart data:', e);
    }
  }

  /**
   * Parse CSV text to chart data
   */
  private parseCSVToChartData(csv: string): any {
    const lines = csv.trim().split('\n');
    const headers = lines[0].split(',').map(h => h.trim());
    const labels: string[] = [];
    const datasets: any[] = [];

    // Initialize datasets for each column (except first which is labels)
    for (let i = 1; i < headers.length; i++) {
      datasets.push({
        label: headers[i],
        data: [],
      });
    }

    // Parse data rows
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(',').map(v => v.trim());
      labels.push(values[0]);
      
      for (let j = 1; j < values.length; j++) {
        const num = parseFloat(values[j]);
        if (!isNaN(num)) {
          datasets[j - 1].data.push(num);
        }
      }
    }

    return {
      labels,
      datasets,
    };
  }

  /**
   * Parse HTML table row
   */
  private parseHTMLTableRow(html: string, cellTag: 'th' | 'td'): string[] {
    // Extract cell contents from HTML
    const regex = new RegExp(`<${cellTag}[^>]*>([^<]*)<\/${cellTag}>`, 'g');
    const cells: string[] = [];
    let match;
    
    while ((match = regex.exec(html)) !== null) {
      cells.push(this.cleanText(match[1]));
    }
    
    return cells;
  }

  /**
   * Check if table contains chart data
   */
  private isChartTable(headers: string[], rows: string[][]): boolean {
    if (headers.length < 2 || rows.length < 2) return false;

    // Check if most cells in data columns are numeric
    let numericCount = 0;
    let totalCount = 0;

    for (const row of rows) {
      for (let i = 1; i < row.length; i++) {
        totalCount++;
        if (!isNaN(parseFloat(row[i]))) {
          numericCount++;
        }
      }
    }

    return numericCount / totalCount > 0.8; // 80% numeric threshold
  }

  /**
   * Create chart slide from table data
   */
  private createChartFromTable(headers: string[], rows: string[][], context: ParsingContext): ChartSlide {
    const labels = rows.map(row => row[0]);
    const datasets = [];

    for (let i = 1; i < headers.length; i++) {
      datasets.push({
        label: headers[i],
        data: rows.map(row => parseFloat(row[i]) || 0),
      });
    }

    return {
      layout: 'chart',
      title: context.currentSlide?.title || 'Data Chart',
      chartType: 'bar',
      data: {
        labels,
        datasets,
      },
    };
  }

  /**
   * Detect indent level for nested lists
   */
  private detectIndentLevel(text: string): number {
    const leadingSpaces = text.match(/^(\s*)/);
    if (leadingSpaces) {
      return Math.floor(leadingSpaces[1].length / 2);
    }
    return 0;
  }

  /**
   * Parse theme configuration
   */
  private parseTheme(theme: string | any): any {
    if (typeof theme === 'string') {
      // Predefined themes
      const themes: Record<string, any> = {
        professional: {
          primaryColor: '#2C3E50',
          secondaryColor: '#34495E',
          fontFamily: 'Arial',
        },
        dark: {
          primaryColor: '#1A1A1A',
          secondaryColor: '#2C2C2C',
          fontFamily: 'Helvetica',
        },
        light: {
          primaryColor: '#FFFFFF',
          secondaryColor: '#F5F5F5',
          fontFamily: 'Calibri',
        },
        academic: {
          primaryColor: '#003366',
          secondaryColor: '#005599',
          fontFamily: 'Times New Roman',
        },
      };

      return themes[theme] || themes.professional;
    }

    return theme;
  }

  /**
   * Insert table of contents slide
   */
  private insertTableOfContents(presentation: Presentation): void {
    const tocBullets: string[] = [];
    
    // Collect section titles (H2 level)
    for (const slide of presentation.slides) {
      if (slide.layout === 'title' && (slide as TitleSlide).backgroundColor) {
        tocBullets.push((slide as TitleSlide).title);
      }
    }

    if (tocBullets.length > 0) {
      const tocSlide: TextSlide = {
        layout: 'text',
        title: 'Table of Contents',
        bullets: tocBullets,
      };

      // Insert after first title slide
      let insertIndex = 1;
      for (let i = 0; i < presentation.slides.length; i++) {
        if (presentation.slides[i].layout === 'title') {
          insertIndex = i + 1;
          break;
        }
      }

      presentation.slides.splice(insertIndex, 0, tocSlide);
    }
  }

  /**
   * Add slide numbers to all slides
   */
  private addSlideNumbers(presentation: Presentation): void {
    // This would need to be implemented in the renderer
    // For now, we'll add it as a note to track the requirement
    presentation.slides.forEach((slide, index) => {
      if (slide.layout !== 'title') {
        const slideNumber = `${index + 1} / ${presentation.slides.length}`;
        // This would need renderer support to actually display
        (slide as any).slideNumber = slideNumber;
      }
    });
  }
}

/**
 * Simplified API function for direct markdown to PowerPoint conversion
 */
export async function markdownToPowerPoint(
  markdownContent: string,
  options?: MarkdownConverterOptions
): Promise<Presentation> {
  const converter = new MarkdownConverter(options);
  return converter.convert(markdownContent);
}