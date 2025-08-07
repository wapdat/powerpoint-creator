/**
 * PDF converter module
 * Handles conversion of PPTX files to PDF using various methods
 */

import { exec } from 'child_process';
import { promisify } from 'util';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

const execAsync = promisify(exec);

/**
 * PDF conversion options
 */
export interface PdfConversionOptions {
  inputPath: string;
  outputPath?: string;
  method?: 'libreoffice' | 'puppeteer' | 'native';
  quality?: 'low' | 'medium' | 'high';
  timeout?: number;
}

/**
 * PDF converter class
 */
export class PdfConverter {
  private converters: Map<string, () => Promise<boolean>>;
  
  constructor() {
    this.converters = new Map();
    this.initializeConverters();
  }
  
  /**
   * Initialize available converters
   */
  private initializeConverters(): void {
    // Check for LibreOffice
    this.converters.set('libreoffice', async () => {
      try {
        await execAsync('libreoffice --version');
        return true;
      } catch {
        try {
          await execAsync('soffice --version');
          return true;
        } catch {
          return false;
        }
      }
    });
    
    // Check for Microsoft PowerPoint (Windows/Mac)
    this.converters.set('powerpoint', async () => {
      if (os.platform() === 'win32') {
        try {
          await execAsync('powershell Get-Command powerpnt.exe');
          return true;
        } catch {
          return false;
        }
      } else if (os.platform() === 'darwin') {
        try {
          await execAsync('osascript -e \'application "Microsoft PowerPoint"\'');
          return true;
        } catch {
          return false;
        }
      }
      return false;
    });
  }
  
  /**
   * Convert PPTX to PDF
   */
  async convert(options: PdfConversionOptions): Promise<void> {
    // Validate input file
    if (!fs.existsSync(options.inputPath)) {
      throw new Error(`Input file not found: ${options.inputPath}`);
    }
    
    if (!options.inputPath.endsWith('.pptx')) {
      throw new Error('Input file must be a .pptx file');
    }
    
    // Set default output path
    const outputPath = options.outputPath || options.inputPath.replace(/\.pptx$/, '.pdf');
    
    // Try conversion methods based on availability and preference
    const method = options.method || 'libreoffice';
    
    try {
      switch (method) {
        case 'libreoffice':
          await this.convertWithLibreOffice(options.inputPath, outputPath, options);
          break;
        case 'puppeteer':
          await this.convertWithPuppeteer(options.inputPath, outputPath, options);
          break;
        case 'native':
          await this.convertWithNative(options.inputPath, outputPath, options);
          break;
        default:
          throw new Error(`Unknown conversion method: ${method}`);
      }
      
      console.log(`PDF generated successfully: ${outputPath}`);
    } catch (error) {
      console.error(`Failed to convert to PDF using ${method}:`, error);
      
      // Try fallback methods
      if (method !== 'libreoffice') {
        console.log('Trying LibreOffice as fallback...');
        try {
          await this.convertWithLibreOffice(options.inputPath, outputPath, options);
          console.log(`PDF generated successfully with LibreOffice: ${outputPath}`);
          return;
        } catch (fallbackError) {
          console.error('LibreOffice fallback failed:', fallbackError);
        }
      }
      
      throw new Error(`PDF conversion failed: ${(error as Error).message}`);
    }
  }
  
  /**
   * Convert using LibreOffice
   */
  private async convertWithLibreOffice(
    inputPath: string,
    outputPath: string,
    options: PdfConversionOptions
  ): Promise<void> {
    // Check if LibreOffice is available
    const isAvailable = await this.converters.get('libreoffice')?.();
    if (!isAvailable) {
      throw new Error('LibreOffice is not installed. Please install LibreOffice to enable PDF conversion.');
    }
    
    // Prepare command
    const outputDir = path.dirname(outputPath);
    const outputFilename = path.basename(outputPath, '.pdf');
    const tempOutputPath = path.join(outputDir, `${outputFilename}.pdf`);
    
    // LibreOffice command with quality settings
    let qualityArgs = '';
    switch (options.quality) {
      case 'high':
        qualityArgs = '--convert-to pdf:writer_pdf_Export:{"MaxImageResolution":{"type":"long","value":"300"}}';
        break;
      case 'low':
        qualityArgs = '--convert-to pdf:writer_pdf_Export:{"MaxImageResolution":{"type":"long","value":"75"}}';
        break;
      default:
        qualityArgs = '--convert-to pdf';
    }
    
    const command = `libreoffice --headless --invisible --nodefault --nolockcheck --nologo --norestore ${qualityArgs} --outdir "${outputDir}" "${inputPath}"`;
    
    try {
      // Execute conversion
      const timeout = options.timeout || 30000;
      const { stdout, stderr } = await execAsync(command, { timeout });
      
      if (stderr && !stderr.includes('Warning')) {
        console.warn('LibreOffice conversion warnings:', stderr);
      }
      
      // Check if output file was created
      const expectedOutput = path.join(outputDir, path.basename(inputPath).replace('.pptx', '.pdf'));
      if (fs.existsSync(expectedOutput)) {
        // Rename to desired output path if different
        if (expectedOutput !== outputPath) {
          fs.renameSync(expectedOutput, outputPath);
        }
      } else {
        throw new Error('PDF file was not created');
      }
      
    } catch (error) {
      if ((error as any).code === 'ETIMEDOUT') {
        throw new Error(`Conversion timed out after ${timeout}ms`);
      }
      throw error;
    }
  }
  
  /**
   * Convert using Puppeteer (requires additional setup)
   */
  private async convertWithPuppeteer(
    inputPath: string,
    outputPath: string,
    options: PdfConversionOptions
  ): Promise<void> {
    // This would require:
    // 1. Converting PPTX to HTML first
    // 2. Using Puppeteer to render HTML and save as PDF
    // This is a placeholder for future implementation
    
    throw new Error('Puppeteer conversion is not yet implemented. Please use LibreOffice method.');
  }
  
  /**
   * Convert using native PowerPoint (Windows/Mac)
   */
  private async convertWithNative(
    inputPath: string,
    outputPath: string,
    options: PdfConversionOptions
  ): Promise<void> {
    const platform = os.platform();
    
    if (platform === 'win32') {
      await this.convertWithPowerPointWindows(inputPath, outputPath);
    } else if (platform === 'darwin') {
      await this.convertWithPowerPointMac(inputPath, outputPath);
    } else {
      throw new Error('Native conversion is only available on Windows and macOS');
    }
  }
  
  /**
   * Convert using PowerPoint on Windows
   */
  private async convertWithPowerPointWindows(
    inputPath: string,
    outputPath: string
  ): Promise<void> {
    const script = `
      $ppt = New-Object -ComObject PowerPoint.Application
      $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
      $presentation = $ppt.Presentations.Open("${inputPath.replace(/\\/g, '\\\\')}")
      $presentation.SaveAs("${outputPath.replace(/\\/g, '\\\\')}", 32) # 32 = ppSaveAsPDF
      $presentation.Close()
      $ppt.Quit()
    `;
    
    try {
      await execAsync(`powershell -Command "${script}"`);
    } catch (error) {
      throw new Error(`PowerPoint conversion failed: ${(error as Error).message}`);
    }
  }
  
  /**
   * Convert using PowerPoint on macOS
   */
  private async convertWithPowerPointMac(
    inputPath: string,
    outputPath: string
  ): Promise<void> {
    const script = `
      tell application "Microsoft PowerPoint"
        open "${inputPath}"
        save active presentation in "${outputPath}" as save as PDF
        close active presentation
      end tell
    `;
    
    try {
      await execAsync(`osascript -e '${script}'`);
    } catch (error) {
      throw new Error(`PowerPoint conversion failed: ${(error as Error).message}`);
    }
  }
  
  /**
   * Check available conversion methods
   */
  async getAvailableMethods(): Promise<string[]> {
    const available: string[] = [];
    
    for (const [method, checker] of this.converters) {
      if (await checker()) {
        available.push(method);
      }
    }
    
    return available;
  }
  
  /**
   * Install LibreOffice instructions
   */
  static getInstallInstructions(): string {
    const platform = os.platform();
    
    switch (platform) {
      case 'darwin':
        return `
To install LibreOffice on macOS:
1. Using Homebrew: brew install --cask libreoffice
2. Or download from: https://www.libreoffice.org/download/download/
        `;
      case 'win32':
        return `
To install LibreOffice on Windows:
1. Download from: https://www.libreoffice.org/download/download/
2. Run the installer and follow the instructions
        `;
      case 'linux':
        return `
To install LibreOffice on Linux:
- Ubuntu/Debian: sudo apt-get install libreoffice
- Fedora: sudo dnf install libreoffice
- Arch: sudo pacman -S libreoffice
- Or download from: https://www.libreoffice.org/download/download/
        `;
      default:
        return 'Please visit https://www.libreoffice.org/download/download/ to install LibreOffice';
    }
  }
}