/**
 * Template processor module using pptx-automizer
 * Handles template-based presentation generation with placeholder replacement
 */

import Automizer, { CmMode, modify, ModifyShapeHelper, ModifyTableHelper } from 'pptx-automizer';
import * as fs from 'fs';
import * as path from 'path';
import {
  Presentation,
  Slide,
  TitleSlide,
  TextSlide,
  ImageSlide,
  ChartSlide,
  TableSlide,
  ChartData,
} from './types';

/**
 * Template processor for working with existing PPTX templates
 */
export class TemplateProcessor {
  private automizer: Automizer;
  
  constructor() {
    this.automizer = new Automizer({
      templateDir: path.join(process.cwd(), 'templates'),
      outputDir: path.join(process.cwd(), 'output'),
      mediaDir: path.join(process.cwd(), 'media'),
      removeExistingSlides: false,
      cleanup: true,
    });
  }
  
  /**
   * Process presentation with template
   */
  async processWithTemplate(
    presentation: Presentation,
    templatePath: string
  ): Promise<Buffer> {
    // Validate template exists
    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template file not found: ${templatePath}`);
    }
    
    // Load template
    const pres = this.automizer.loadRoot(templatePath);
    
    // Get slide masters from template
    const templateInfo = await this.inspectTemplate(templatePath);
    
    // Process each slide
    for (let i = 0; i < presentation.slides.length; i++) {
      const slideData = presentation.slides[i];
      await this.processSlide(pres, slideData, i, templateInfo);
    }
    
    // Generate output
    const result = await this.automizer.write(
      `presentation_${Date.now()}.pptx`
    );
    
    // Read the generated file and return as buffer
    const outputPath = result.files[0];
    const buffer = fs.readFileSync(outputPath);
    
    // Clean up temporary file
    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
    
    return buffer;
  }
  
  /**
   * Inspect template to get available layouts and placeholders
   */
  async inspectTemplate(templatePath: string): Promise<any> {
    try {
      const pres = this.automizer.loadRoot(templatePath);
      
      // Get slide count and layouts
      const info = await pres.getInfo();
      
      return {
        slideCount: info.slideCount,
        layouts: info.slideLayouts || [],
        masters: info.slideMasters || [],
        placeholders: await this.extractPlaceholders(pres),
      };
    } catch (error) {
      throw new Error(`Failed to inspect template: ${(error as Error).message}`);
    }
  }
  
  /**
   * Extract placeholders from template
   */
  private async extractPlaceholders(pres: any): Promise<any[]> {
    const placeholders: any[] = [];
    
    try {
      // Get first slide to analyze placeholders
      const slideInfo = await pres.getSlide(1).getInfo();
      
      if (slideInfo.shapes) {
        slideInfo.shapes.forEach((shape: any) => {
          if (shape.type === 'placeholder') {
            placeholders.push({
              name: shape.name,
              type: shape.placeholderType,
              id: shape.id,
            });
          }
        });
      }
    } catch (error) {
      console.warn('Could not extract placeholders:', error);
    }
    
    return placeholders;
  }
  
  /**
   * Process individual slide
   */
  private async processSlide(
    pres: any,
    slideData: Slide,
    index: number,
    templateInfo: any
  ): Promise<void> {
    // Clone appropriate template slide or add new slide
    const slideNumber = index + 1;
    
    // If we're beyond template slides, clone the last one
    if (slideNumber > templateInfo.slideCount) {
      await pres.addSlide(templateInfo.slideCount);
    }
    
    // Get the slide to modify
    const slide = pres.getSlide(slideNumber);
    
    // Process based on layout type
    switch (slideData.layout) {
      case 'title':
        await this.processTitleSlide(slide, slideData as TitleSlide);
        break;
      case 'text':
        await this.processTextSlide(slide, slideData as TextSlide);
        break;
      case 'image':
        await this.processImageSlide(slide, slideData as ImageSlide);
        break;
      case 'chart':
        await this.processChartSlide(slide, slideData as ChartSlide);
        break;
      case 'table':
        await this.processTableSlide(slide, slideData as TableSlide);
        break;
      default:
        await this.processGenericSlide(slide, slideData);
    }
    
    // Add speaker notes if provided
    if (slideData.notes) {
      await this.addSpeakerNotes(slide, slideData.notes);
    }
  }
  
  /**
   * Process title slide
   */
  private async processTitleSlide(slide: any, slideData: TitleSlide): Promise<void> {
    // Replace title placeholder
    await slide.modifyElement(
      ModifyShapeHelper.setText({
        type: 'placeholder',
        placeholderType: 'title',
      }, slideData.title)
    );
    
    // Replace subtitle placeholder
    if (slideData.subtitle) {
      await slide.modifyElement(
        ModifyShapeHelper.setText({
          type: 'placeholder',
          placeholderType: 'subtitle',
        }, slideData.subtitle)
      );
    }
    
    // Add author and date as footer text
    if (slideData.author || slideData.date) {
      const footerText = [slideData.author, slideData.date].filter(Boolean).join(' | ');
      await slide.modifyElement(
        ModifyShapeHelper.setText({
          type: 'placeholder',
          placeholderType: 'footer',
        }, footerText)
      );
    }
  }
  
  /**
   * Process text slide with bullets
   */
  private async processTextSlide(slide: any, slideData: TextSlide): Promise<void> {
    // Replace title
    await slide.modifyElement(
      ModifyShapeHelper.setText({
        type: 'placeholder',
        placeholderType: 'title',
      }, slideData.title)
    );
    
    // Format bullets
    const bulletText = slideData.bullets.map((bullet, index) => {
      const level = slideData.level?.[index] || 0;
      const indent = '  '.repeat(level);
      return `${indent}• ${bullet}`;
    }).join('\n');
    
    // Replace body placeholder with bullets
    await slide.modifyElement(
      ModifyShapeHelper.setText({
        type: 'placeholder',
        placeholderType: 'body',
      }, bulletText)
    );
  }
  
  /**
   * Process image slide
   */
  private async processImageSlide(slide: any, slideData: ImageSlide): Promise<void> {
    // Replace title
    await slide.modifyElement(
      ModifyShapeHelper.setText({
        type: 'placeholder',
        placeholderType: 'title',
      }, slideData.title)
    );
    
    // Replace image placeholder
    if (slideData.imagePath || slideData.imageUrl) {
      let imagePath = slideData.imagePath;
      
      // Download image if URL provided
      if (slideData.imageUrl) {
        imagePath = await this.downloadImage(slideData.imageUrl);
      }
      
      if (imagePath && fs.existsSync(imagePath)) {
        await slide.modifyElement(
          modify.setImage({
            type: 'placeholder',
            placeholderType: 'picture',
          }, imagePath)
        );
      }
    }
    
    // Add caption if provided
    if (slideData.caption) {
      await slide.modifyElement(
        ModifyShapeHelper.setText({
          type: 'placeholder',
          placeholderType: 'body',
        }, slideData.caption)
      );
    }
  }
  
  /**
   * Process chart slide
   */
  private async processChartSlide(slide: any, slideData: ChartSlide): Promise<void> {
    // Replace title
    await slide.modifyElement(
      ModifyShapeHelper.setText({
        type: 'placeholder',
        placeholderType: 'title',
      }, slideData.title)
    );
    
    // Convert chart data to automizer format
    const chartData = this.convertChartData(slideData.data, slideData.chartType);
    
    // Replace chart placeholder
    await slide.modifyElement(
      modify.setChartData({
        type: 'chart',
      }, chartData)
    );
  }
  
  /**
   * Process table slide
   */
  private async processTableSlide(slide: any, slideData: TableSlide): Promise<void> {
    // Replace title
    await slide.modifyElement(
      ModifyShapeHelper.setText({
        type: 'placeholder',
        placeholderType: 'title',
      }, slideData.title)
    );
    
    // Prepare table data
    const tableData = [];
    
    // Add headers if provided
    if (slideData.headers) {
      tableData.push(slideData.headers);
    }
    
    // Add data rows
    tableData.push(...slideData.tableData);
    
    // Replace table placeholder
    await slide.modifyElement(
      ModifyTableHelper.setData({
        type: 'table',
      }, tableData)
    );
  }
  
  /**
   * Process generic slide
   */
  private async processGenericSlide(slide: any, slideData: Slide): Promise<void> {
    // Replace title if present
    if (slideData.title) {
      await slide.modifyElement(
        ModifyShapeHelper.setText({
          type: 'placeholder',
          placeholderType: 'title',
        }, slideData.title)
      );
    }
    
    // Replace subtitle if present
    if (slideData.subtitle) {
      await slide.modifyElement(
        ModifyShapeHelper.setText({
          type: 'placeholder',
          placeholderType: 'subtitle',
        }, slideData.subtitle)
      );
    }
  }
  
  /**
   * Add speaker notes to slide
   */
  private async addSpeakerNotes(slide: any, notes: string): Promise<void> {
    try {
      await slide.modifyNotes(notes);
    } catch (error) {
      console.warn('Could not add speaker notes:', error);
    }
  }
  
  /**
   * Convert chart data to automizer format
   */
  private convertChartData(data: ChartData, chartType: string): any {
    const series = data.datasets.map(dataset => ({
      name: dataset.label,
      values: dataset.data,
    }));
    
    return {
      categories: data.labels,
      series: series,
      chartType: chartType,
    };
  }
  
  /**
   * Download image from URL
   */
  private async downloadImage(url: string): Promise<string | null> {
    try {
      const axios = require('axios');
      const response = await axios.get(url, {
        responseType: 'arraybuffer',
      });
      
      // Save to temporary file
      const tempPath = path.join(
        process.cwd(),
        'temp',
        `image_${Date.now()}.jpg`
      );
      
      // Ensure temp directory exists
      const tempDir = path.dirname(tempPath);
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }
      
      fs.writeFileSync(tempPath, response.data);
      
      return tempPath;
    } catch (error) {
      console.error(`Failed to download image: ${url}`, error);
      return null;
    }
  }
  
  /**
   * Clean up temporary files
   */
  private cleanupTempFiles(): void {
    const tempDir = path.join(process.cwd(), 'temp');
    if (fs.existsSync(tempDir)) {
      const files = fs.readdirSync(tempDir);
      files.forEach(file => {
        const filePath = path.join(tempDir, file);
        if (fs.existsSync(filePath)) {
          fs.unlinkSync(filePath);
        }
      });
    }
  }
}