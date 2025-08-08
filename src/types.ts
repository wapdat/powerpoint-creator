/**
 * Type definitions for pptx-auto-gen
 * Defines the structure for presentation input and internal data models
 */

export type SlideLayout = 'title' | 'text' | 'image' | 'chart' | 'table' | 'notes' | 'custom';

export type ChartType = 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'doughnut' | 'radar';

/**
 * Base slide interface with common properties
 */
export interface BaseSlide {
  layout: SlideLayout;
  title?: string;
  subtitle?: string;
  notes?: string;
  backgroundColor?: string;
  transition?: 'none' | 'fade' | 'slide' | 'convex' | 'concave' | 'zoom';
}

/**
 * Title slide with main title and subtitle
 */
export interface TitleSlide extends BaseSlide {
  layout: 'title';
  title: string;
  subtitle?: string;
  author?: string;
  date?: string;
}

/**
 * Text slide with bullet points
 */
export interface TextSlide extends BaseSlide {
  layout: 'text';
  title: string;
  bullets: string[];
  level?: number[];  // Indentation levels for bullets
}

/**
 * Image slide with image content
 */
export interface ImageSlide extends BaseSlide {
  layout: 'image';
  title: string;
  imagePath?: string;
  imageUrl?: string;
  imageAlt?: string;
  caption?: string;
  sizing?: 'contain' | 'cover' | 'stretch';
}

/**
 * Chart slide with data visualization
 */
export interface ChartSlide extends BaseSlide {
  layout: 'chart';
  title: string;
  chartType: ChartType;
  data: ChartData;
  options?: ChartOptions;
}

/**
 * Table slide with tabular data
 */
export interface TableSlide extends BaseSlide {
  layout: 'table';
  title: string;
  tableData: (string | number)[][];
  headers?: string[];
  styling?: TableStyling;
}

/**
 * Notes slide for speaker notes only
 */
export interface NotesSlide extends BaseSlide {
  layout: 'notes';
  title: string;
  content: string;
}

/**
 * Custom slide for advanced layouts
 */
export interface CustomSlide extends BaseSlide {
  layout: 'custom';
  title: string;
  elements: SlideElement[];
}

/**
 * Union type for all slide types
 */
export type Slide = TitleSlide | TextSlide | ImageSlide | ChartSlide | TableSlide | NotesSlide | CustomSlide;

/**
 * Chart data structure
 */
export interface ChartData {
  labels: string[];
  datasets: {
    label: string;
    data: number[];
    backgroundColor?: string | string[];
    borderColor?: string | string[];
    borderWidth?: number;
  }[];
}

/**
 * Chart options for customization
 */
export interface ChartOptions {
  showLegend?: boolean;
  showTitle?: boolean;
  showDataLabels?: boolean;
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  colors?: string[];
}

/**
 * Table styling options
 */
export interface TableStyling {
  headerBackground?: string;
  headerTextColor?: string;
  borderColor?: string;
  borderWidth?: number;
  alternateRows?: boolean;
  fontSize?: number;
}

/**
 * Custom slide elements
 */
export interface SlideElement {
  type: 'text' | 'image' | 'shape' | 'chart' | 'table';
  x: number;  // Position in percentage or pixels
  y: number;
  width: number;
  height: number;
  content: any;
  styling?: any;
}

/**
 * Main presentation structure
 */
export interface Presentation {
  title: string;
  author?: string;
  subject?: string;
  company?: string;
  slides: Slide[];
  theme?: PresentationTheme;
  metadata?: PresentationMetadata;
}

/**
 * Presentation theme settings
 */
export interface PresentationTheme {
  primaryColor?: string;
  secondaryColor?: string;
  fontFamily?: string;
  fontSize?: number;
  backgroundColor?: string;
  accentColor?: string;
}

/**
 * Presentation metadata
 */
export interface PresentationMetadata {
  created?: Date;
  modified?: Date;
  revision?: string;
  keywords?: string[];
  category?: string;
}

/**
 * CLI options interface
 */
export interface CLIOptions {
  input?: string;
  markdown?: string;
  output: string;
  template?: string;
  pdf?: boolean;
  verbose?: boolean;
  help?: boolean;
  version?: boolean;
}

/**
 * Generation options for the main API
 */
export interface GenerationOptions {
  inputPath?: string;
  inputData?: Presentation;
  outputPath: string;
  templatePath?: string;
  convertToPdf?: boolean;
  validation?: boolean;
  styling?: PresentationTheme;
}

/**
 * Validation result interface
 */
export interface ValidationResult {
  valid: boolean;
  errors?: ValidationError[];
}

/**
 * Validation error details
 */
export interface ValidationError {
  field: string;
  message: string;
  value?: any;
}

/**
 * PDF conversion options
 */
export interface PdfConversionOptions {
  method?: 'libreoffice' | 'puppeteer' | 'native';
  quality?: 'low' | 'medium' | 'high';
  outputPath?: string;
}