/**
 * Main API module for pptx-auto-gen
 * Provides programmatic interface for presentation generation
 */

import * as fs from 'fs';
import * as path from 'path';
import { GenerationOptions, Presentation, ValidationResult } from './types';
import { SlideRenderer } from './renderer';
import { TemplateProcessor } from './template';
import { PdfConverter } from './pdf-converter';
import { InputValidator } from './validator';

/**
 * Main function to generate PowerPoint presentations
 * @param options Generation options including input data, output path, and optional template
 * @returns Promise that resolves when generation is complete
 */
export async function generatePresentation(options: GenerationOptions): Promise<void> {
  try {
    // Load input data if path is provided
    let presentationData: Presentation;
    
    if (options.inputPath) {
      const inputContent = fs.readFileSync(options.inputPath, 'utf-8');
      presentationData = JSON.parse(inputContent);
    } else if (options.inputData) {
      presentationData = options.inputData;
    } else {
      throw new Error('Either inputPath or inputData must be provided');
    }
    
    // Validate input if requested
    if (options.validation !== false) {
      const validator = new InputValidator();
      const validationResult = validator.validate(presentationData);
      
      if (!validationResult.valid) {
        const errorMessages = validationResult.errors?.map(e => `${e.field}: ${e.message}`).join('\n');
        throw new Error(`Validation failed:\n${errorMessages}`);
      }
    }
    
    // Apply styling overrides if provided
    if (options.styling) {
      presentationData.theme = {
        ...presentationData.theme,
        ...options.styling,
      };
    }
    
    // Generate presentation based on template availability
    let outputBuffer: Buffer;
    
    if (options.templatePath) {
      // Use template processor for template-based generation
      const templateProcessor = new TemplateProcessor();
      outputBuffer = await templateProcessor.processWithTemplate(
        presentationData,
        options.templatePath
      );
    } else {
      // Use renderer for generation from scratch
      const renderer = new SlideRenderer();
      outputBuffer = await renderer.renderPresentation(presentationData);
    }
    
    // Ensure output directory exists
    const outputDir = path.dirname(options.outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    // Write the presentation file
    fs.writeFileSync(options.outputPath, outputBuffer);
    
    // Convert to PDF if requested
    if (options.convertToPdf) {
      const pdfConverter = new PdfConverter();
      const pdfPath = options.outputPath.replace(/\.pptx$/, '.pdf');
      
      await pdfConverter.convert({
        inputPath: options.outputPath,
        outputPath: pdfPath,
        method: 'libreoffice',
        quality: 'high',
      });
    }
    
  } catch (error) {
    throw new Error(`Failed to generate presentation: ${(error as Error).message}`);
  }
}

/**
 * Validate presentation input without generating output
 * @param input Presentation data to validate
 * @returns Validation result with any errors
 */
export function validateInput(input: Presentation): ValidationResult {
  const validator = new InputValidator();
  return validator.validate(input);
}

/**
 * Load a presentation template for inspection
 * @param templatePath Path to the template file
 * @returns Template information and available layouts
 */
export async function inspectTemplate(templatePath: string): Promise<any> {
  const templateProcessor = new TemplateProcessor();
  return templateProcessor.inspectTemplate(templatePath);
}

/**
 * Get available slide layouts
 * @returns List of supported slide layouts
 */
export function getAvailableLayouts(): string[] {
  return ['title', 'text', 'image', 'chart', 'table', 'notes', 'custom'];
}

/**
 * Get available chart types
 * @returns List of supported chart types
 */
export function getAvailableChartTypes(): string[] {
  return ['bar', 'line', 'pie', 'area', 'scatter', 'doughnut', 'radar'];
}

/**
 * Create a sample presentation structure
 * @param slideCount Number of slides to include
 * @returns Sample presentation object
 */
export function createSamplePresentation(slideCount: number = 5): Presentation {
  const slides = [];
  
  // Add title slide
  slides.push({
    layout: 'title',
    title: 'Sample Presentation',
    subtitle: 'Generated with pptx-auto-gen',
    author: 'Your Name',
    date: new Date().toLocaleDateString(),
  });
  
  // Add text slide
  if (slideCount > 1) {
    slides.push({
      layout: 'text',
      title: 'Key Points',
      bullets: [
        'First important point',
        'Second important point',
        'Third important point with sub-items',
        '  • Sub-item one',
        '  • Sub-item two',
      ],
    });
  }
  
  // Add chart slide
  if (slideCount > 2) {
    slides.push({
      layout: 'chart',
      title: 'Performance Metrics',
      chartType: 'bar',
      data: {
        labels: ['Q1', 'Q2', 'Q3', 'Q4'],
        datasets: [
          {
            label: 'Revenue',
            data: [65, 75, 85, 95],
            backgroundColor: '#4472C4',
          },
          {
            label: 'Profit',
            data: [28, 35, 40, 48],
            backgroundColor: '#ED7D31',
          },
        ],
      },
    });
  }
  
  // Add table slide
  if (slideCount > 3) {
    slides.push({
      layout: 'table',
      title: 'Quarterly Results',
      headers: ['Quarter', 'Revenue', 'Profit', 'Growth'],
      tableData: [
        ['Q1 2024', '$1.2M', '$280K', '15%'],
        ['Q2 2024', '$1.5M', '$350K', '25%'],
        ['Q3 2024', '$1.8M', '$400K', '20%'],
        ['Q4 2024', '$2.1M', '$480K', '17%'],
      ],
    });
  }
  
  // Add image slide
  if (slideCount > 4) {
    slides.push({
      layout: 'image',
      title: 'Product Overview',
      imageUrl: 'https://via.placeholder.com/800x600',
      caption: 'Our flagship product',
    });
  }
  
  return {
    title: 'Sample Presentation',
    author: 'pptx-auto-gen',
    slides: slides as any,
    theme: {
      primaryColor: '#4472C4',
      secondaryColor: '#ED7D31',
      fontFamily: 'Arial',
      fontSize: 14,
    },
  };
}

// Export all types for consumer convenience
export * from './types';

// Export individual modules for advanced usage
export { SlideRenderer } from './renderer';
export { TemplateProcessor } from './template';
export { PdfConverter } from './pdf-converter';
export { InputValidator } from './validator';