/**
 * Input validator module
 * Validates presentation JSON against schema and business rules
 */

import Ajv from 'ajv';
import {
  Presentation,
  Slide,
  ValidationResult,
  ValidationError,
  SlideLayout,
  ChartType,
} from './types';

/**
 * JSON Schema for presentation validation
 */
const PRESENTATION_SCHEMA = {
  type: 'object',
  required: ['slides'],
  properties: {
    title: { type: 'string' },
    author: { type: 'string' },
    subject: { type: 'string' },
    company: { type: 'string' },
    slides: {
      type: 'array',
      minItems: 1,
      items: {
        type: 'object',
        required: ['layout'],
        properties: {
          layout: {
            type: 'string',
            enum: ['title', 'text', 'image', 'chart', 'table', 'notes', 'custom'],
          },
          title: { type: 'string' },
          subtitle: { type: 'string' },
          notes: { type: 'string' },
          backgroundColor: { type: 'string' },
          transition: {
            type: 'string',
            enum: ['none', 'fade', 'slide', 'convex', 'concave', 'zoom'],
          },
        },
      },
    },
    theme: {
      type: 'object',
      properties: {
        primaryColor: { type: 'string' },
        secondaryColor: { type: 'string' },
        fontFamily: { type: 'string' },
        fontSize: { type: 'number', minimum: 8, maximum: 72 },
        backgroundColor: { type: 'string' },
        accentColor: { type: 'string' },
      },
    },
    metadata: {
      type: 'object',
      properties: {
        created: { type: 'string' },
        modified: { type: 'string' },
        revision: { type: 'string' },
        keywords: {
          type: 'array',
          items: { type: 'string' },
        },
        category: { type: 'string' },
      },
    },
  },
};

/**
 * Input validator class
 */
export class InputValidator {
  private ajv: Ajv;
  private validateSchema: any;
  
  constructor() {
    this.ajv = new Ajv({ allErrors: true });
    this.validateSchema = this.ajv.compile(PRESENTATION_SCHEMA);
  }
  
  /**
   * Validate presentation input
   */
  validate(presentation: Presentation): ValidationResult {
    const errors: ValidationError[] = [];
    
    // Validate against JSON schema
    const schemaValid = this.validateSchema(presentation);
    if (!schemaValid && this.validateSchema.errors) {
      this.validateSchema.errors.forEach((error: any) => {
        errors.push({
          field: error.instancePath || error.dataPath || 'root',
          message: error.message || 'Invalid value',
          value: error.data,
        });
      });
    }
    
    // Validate individual slides
    if (presentation.slides) {
      presentation.slides.forEach((slide, index) => {
        const slideErrors = this.validateSlide(slide, index);
        errors.push(...slideErrors);
      });
    } else {
      errors.push({
        field: 'slides',
        message: 'Presentation must contain at least one slide',
      });
    }
    
    // Validate theme colors
    if (presentation.theme) {
      const themeErrors = this.validateTheme(presentation.theme);
      errors.push(...themeErrors);
    }
    
    return {
      valid: errors.length === 0,
      errors: errors.length > 0 ? errors : undefined,
    };
  }
  
  /**
   * Validate individual slide
   */
  private validateSlide(slide: Slide, index: number): ValidationError[] {
    const errors: ValidationError[] = [];
    const prefix = `slides[${index}]`;
    
    // Validate layout-specific requirements
    switch (slide.layout) {
      case 'title':
        if (!slide.title) {
          errors.push({
            field: `${prefix}.title`,
            message: 'Title slide must have a title',
            value: slide.title,
          });
        }
        break;
        
      case 'text':
        if (!slide.title) {
          errors.push({
            field: `${prefix}.title`,
            message: 'Text slide must have a title',
            value: slide.title,
          });
        }
        if (!(slide as any).bullets || (slide as any).bullets.length === 0) {
          errors.push({
            field: `${prefix}.bullets`,
            message: 'Text slide must have at least one bullet point',
            value: (slide as any).bullets,
          });
        }
        break;
        
      case 'image':
        if (!slide.title) {
          errors.push({
            field: `${prefix}.title`,
            message: 'Image slide must have a title',
            value: slide.title,
          });
        }
        if (!(slide as any).imagePath && !(slide as any).imageUrl) {
          errors.push({
            field: `${prefix}.image`,
            message: 'Image slide must have either imagePath or imageUrl',
          });
        }
        break;
        
      case 'chart':
        if (!slide.title) {
          errors.push({
            field: `${prefix}.title`,
            message: 'Chart slide must have a title',
            value: slide.title,
          });
        }
        if (!(slide as any).chartType) {
          errors.push({
            field: `${prefix}.chartType`,
            message: 'Chart slide must specify a chart type',
          });
        } else if (!this.isValidChartType((slide as any).chartType)) {
          errors.push({
            field: `${prefix}.chartType`,
            message: `Invalid chart type: ${(slide as any).chartType}`,
            value: (slide as any).chartType,
          });
        }
        if (!(slide as any).data) {
          errors.push({
            field: `${prefix}.data`,
            message: 'Chart slide must have data',
          });
        } else {
          const dataErrors = this.validateChartData((slide as any).data, prefix);
          errors.push(...dataErrors);
        }
        break;
        
      case 'table':
        if (!slide.title) {
          errors.push({
            field: `${prefix}.title`,
            message: 'Table slide must have a title',
            value: slide.title,
          });
        }
        if (!(slide as any).tableData || (slide as any).tableData.length === 0) {
          errors.push({
            field: `${prefix}.tableData`,
            message: 'Table slide must have table data',
            value: (slide as any).tableData,
          });
        } else {
          const tableErrors = this.validateTableData((slide as any).tableData, prefix);
          errors.push(...tableErrors);
        }
        break;
        
      case 'notes':
        if (!slide.title) {
          errors.push({
            field: `${prefix}.title`,
            message: 'Notes slide must have a title',
            value: slide.title,
          });
        }
        if (!(slide as any).content) {
          errors.push({
            field: `${prefix}.content`,
            message: 'Notes slide must have content',
            value: (slide as any).content,
          });
        }
        break;
        
      case 'custom':
        if (!(slide as any).elements || (slide as any).elements.length === 0) {
          errors.push({
            field: `${prefix}.elements`,
            message: 'Custom slide must have at least one element',
            value: (slide as any).elements,
          });
        }
        break;
    }
    
    // Validate colors if present
    if (slide.backgroundColor) {
      if (!this.isValidColor(slide.backgroundColor)) {
        errors.push({
          field: `${prefix}.backgroundColor`,
          message: 'Invalid color format',
          value: slide.backgroundColor,
        });
      }
    }
    
    return errors;
  }
  
  /**
   * Validate chart data
   */
  private validateChartData(data: any, prefix: string): ValidationError[] {
    const errors: ValidationError[] = [];
    
    if (!data.labels || !Array.isArray(data.labels)) {
      errors.push({
        field: `${prefix}.data.labels`,
        message: 'Chart data must have labels array',
        value: data.labels,
      });
    }
    
    if (!data.datasets || !Array.isArray(data.datasets) || data.datasets.length === 0) {
      errors.push({
        field: `${prefix}.data.datasets`,
        message: 'Chart data must have at least one dataset',
        value: data.datasets,
      });
    } else {
      data.datasets.forEach((dataset: any, index: number) => {
        if (!dataset.label) {
          errors.push({
            field: `${prefix}.data.datasets[${index}].label`,
            message: 'Dataset must have a label',
            value: dataset.label,
          });
        }
        if (!dataset.data || !Array.isArray(dataset.data)) {
          errors.push({
            field: `${prefix}.data.datasets[${index}].data`,
            message: 'Dataset must have data array',
            value: dataset.data,
          });
        } else if (data.labels && dataset.data.length !== data.labels.length) {
          errors.push({
            field: `${prefix}.data.datasets[${index}].data`,
            message: `Dataset data length (${dataset.data.length}) must match labels length (${data.labels.length})`,
            value: dataset.data,
          });
        }
      });
    }
    
    return errors;
  }
  
  /**
   * Validate table data
   */
  private validateTableData(tableData: any, prefix: string): ValidationError[] {
    const errors: ValidationError[] = [];
    
    if (!Array.isArray(tableData)) {
      errors.push({
        field: `${prefix}.tableData`,
        message: 'Table data must be a 2D array',
        value: tableData,
      });
      return errors;
    }
    
    let columnCount: number | null = null;
    
    tableData.forEach((row: any, index: number) => {
      if (!Array.isArray(row)) {
        errors.push({
          field: `${prefix}.tableData[${index}]`,
          message: 'Each table row must be an array',
          value: row,
        });
      } else {
        if (columnCount === null) {
          columnCount = row.length;
        } else if (row.length !== columnCount) {
          errors.push({
            field: `${prefix}.tableData[${index}]`,
            message: `Row has ${row.length} columns, expected ${columnCount}`,
            value: row,
          });
        }
      }
    });
    
    return errors;
  }
  
  /**
   * Validate theme
   */
  private validateTheme(theme: any): ValidationError[] {
    const errors: ValidationError[] = [];
    
    const colorFields = [
      'primaryColor',
      'secondaryColor',
      'backgroundColor',
      'accentColor',
    ];
    
    colorFields.forEach(field => {
      if (theme[field] && !this.isValidColor(theme[field])) {
        errors.push({
          field: `theme.${field}`,
          message: 'Invalid color format',
          value: theme[field],
        });
      }
    });
    
    if (theme.fontSize && (theme.fontSize < 8 || theme.fontSize > 72)) {
      errors.push({
        field: 'theme.fontSize',
        message: 'Font size must be between 8 and 72',
        value: theme.fontSize,
      });
    }
    
    return errors;
  }
  
  /**
   * Check if color is valid
   */
  private isValidColor(color: string): boolean {
    // Accept hex colors
    if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
      return true;
    }
    // Accept hex colors with alpha
    if (/^#[0-9A-Fa-f]{8}$/.test(color)) {
      return true;
    }
    // Accept short hex
    if (/^#[0-9A-Fa-f]{3}$/.test(color)) {
      return true;
    }
    // Accept rgb/rgba
    if (/^rgba?\(.*\)$/.test(color)) {
      return true;
    }
    // Accept color names
    const validColorNames = [
      'black', 'white', 'red', 'green', 'blue', 'yellow', 'cyan', 'magenta',
      'gray', 'grey', 'orange', 'purple', 'brown', 'pink', 'lime', 'navy',
      'teal', 'silver', 'gold', 'indigo', 'violet', 'transparent',
    ];
    if (validColorNames.includes(color.toLowerCase())) {
      return true;
    }
    
    return false;
  }
  
  /**
   * Check if chart type is valid
   */
  private isValidChartType(type: string): boolean {
    const validTypes: ChartType[] = [
      'bar', 'line', 'pie', 'area', 'scatter', 'doughnut', 'radar',
    ];
    return validTypes.includes(type as ChartType);
  }
  
  /**
   * Get validation schema
   */
  getSchema(): any {
    return PRESENTATION_SCHEMA;
  }
  
  /**
   * Get sample valid input
   */
  getSampleInput(): Presentation {
    return {
      title: 'Sample Presentation',
      author: 'John Doe',
      slides: [
        {
          layout: 'title',
          title: 'Welcome to My Presentation',
          subtitle: 'A Professional PowerPoint Generator',
          author: 'John Doe',
          date: new Date().toLocaleDateString(),
        } as any,
        {
          layout: 'text',
          title: 'Agenda',
          bullets: [
            'Introduction',
            'Main Points',
            'Conclusion',
          ],
        } as any,
        {
          layout: 'chart',
          title: 'Sales Performance',
          chartType: 'bar',
          data: {
            labels: ['Q1', 'Q2', 'Q3', 'Q4'],
            datasets: [
              {
                label: 'Revenue',
                data: [100, 150, 200, 250],
                backgroundColor: '#4472C4',
              },
            ],
          },
        } as any,
      ],
      theme: {
        primaryColor: '#4472C4',
        secondaryColor: '#ED7D31',
        fontFamily: 'Arial',
        fontSize: 14,
      },
    };
  }
}