#!/usr/bin/env node

/**
 * CLI entry point for pptx-auto-gen
 * Handles command-line arguments and orchestrates presentation generation
 */

import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import * as fs from 'fs';
import * as path from 'path';
import chalk from 'chalk';
import ora from 'ora';
import { generatePresentation } from './index';
import { CLIOptions, Presentation } from './types';

/**
 * Parse command-line arguments
 */
const argv = yargs(hideBin(process.argv))
  .command('$0', 'Generate PowerPoint presentations from JSON input')
  .option('input', {
    alias: 'i',
    type: 'string',
    description: 'Path to JSON input file or use STDIN',
    demandOption: false,
  })
  .option('output', {
    alias: 'o',
    type: 'string',
    description: 'Output PPTX file path',
    default: 'output.pptx',
  })
  .option('template', {
    alias: 't',
    type: 'string',
    description: 'Path to template PPTX file',
  })
  .option('pdf', {
    alias: 'p',
    type: 'boolean',
    description: 'Convert output to PDF',
    default: false,
  })
  .option('verbose', {
    alias: 'v',
    type: 'boolean',
    description: 'Enable verbose logging',
    default: false,
  })
  .help()
  .alias('help', 'h')
  .example(
    '$0 --input slides.json --output presentation.pptx',
    'Generate presentation from JSON file'
  )
  .example(
    '$0 --input slides.json --template corporate.pptx --output final.pptx',
    'Generate using a template'
  )
  .example(
    'cat slides.json | $0 --output presentation.pptx',
    'Generate from STDIN'
  )
  .epilogue('For more information, visit https://github.com/yourusername/pptx-auto-gen')
  .parseSync() as CLIOptions;

/**
 * Main CLI execution function
 */
async function main(): Promise<void> {
  const spinner = ora();
  
  try {
    // Handle input source
    let inputData: Presentation;
    
    if (argv.input) {
      // Read from file
      spinner.start(chalk.blue('Reading input file...'));
      
      if (!fs.existsSync(argv.input)) {
        spinner.fail(chalk.red(`Input file not found: ${argv.input}`));
        process.exit(1);
      }
      
      const inputContent = fs.readFileSync(argv.input, 'utf-8');
      
      try {
        inputData = JSON.parse(inputContent);
        spinner.succeed(chalk.green('Input file parsed successfully'));
      } catch (error) {
        spinner.fail(chalk.red('Invalid JSON in input file'));
        if (argv.verbose) {
          console.error(error);
        }
        process.exit(1);
      }
    } else {
      // Read from STDIN
      spinner.start(chalk.blue('Reading from STDIN...'));
      
      const chunks: string[] = [];
      
      await new Promise<void>((resolve, reject) => {
        process.stdin.setEncoding('utf-8');
        
        process.stdin.on('data', (chunk) => {
          chunks.push(chunk);
        });
        
        process.stdin.on('end', () => {
          resolve();
        });
        
        process.stdin.on('error', (error) => {
          reject(error);
        });
        
        // Set timeout for STDIN
        setTimeout(() => {
          if (chunks.length === 0) {
            reject(new Error('No input received from STDIN'));
          }
        }, 5000);
      });
      
      try {
        inputData = JSON.parse(chunks.join(''));
        spinner.succeed(chalk.green('STDIN input parsed successfully'));
      } catch (error) {
        spinner.fail(chalk.red('Invalid JSON from STDIN'));
        if (argv.verbose) {
          console.error(error);
        }
        process.exit(1);
      }
    }
    
    // Validate template if provided
    if (argv.template) {
      spinner.start(chalk.blue('Validating template file...'));
      
      if (!fs.existsSync(argv.template)) {
        spinner.fail(chalk.red(`Template file not found: ${argv.template}`));
        process.exit(1);
      }
      
      if (!argv.template.endsWith('.pptx')) {
        spinner.fail(chalk.red('Template must be a .pptx file'));
        process.exit(1);
      }
      
      spinner.succeed(chalk.green('Template file validated'));
    }
    
    // Ensure output directory exists
    const outputDir = path.dirname(argv.output);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    // Generate presentation
    spinner.start(chalk.blue('Generating presentation...'));
    
    await generatePresentation({
      inputData,
      outputPath: argv.output,
      templatePath: argv.template,
      convertToPdf: argv.pdf,
      validation: true,
    });
    
    spinner.succeed(chalk.green(`Presentation generated: ${argv.output}`));
    
    // Handle PDF conversion if requested
    if (argv.pdf) {
      spinner.start(chalk.blue('Converting to PDF...'));
      
      const pdfPath = argv.output.replace(/\.pptx$/, '.pdf');
      
      // PDF conversion will be handled in the pdf-converter module
      spinner.info(chalk.yellow(`PDF conversion requested. Output will be: ${pdfPath}`));
    }
    
    // Success message
    console.log(chalk.green.bold('\n✨ Presentation generated successfully!'));
    
    if (argv.verbose) {
      console.log(chalk.gray('\nGeneration details:'));
      console.log(chalk.gray(`  • Slides: ${inputData.slides.length}`));
      console.log(chalk.gray(`  • Template: ${argv.template || 'Default'}`));
      console.log(chalk.gray(`  • Output: ${path.resolve(argv.output)}`));
      
      if (argv.pdf) {
        console.log(chalk.gray(`  • PDF: ${path.resolve(argv.output.replace(/\.pptx$/, '.pdf'))}`));
      }
    }
    
  } catch (error) {
    spinner.fail(chalk.red('Failed to generate presentation'));
    
    if (argv.verbose) {
      console.error(chalk.red('\nError details:'));
      console.error(error);
    } else {
      console.error(chalk.red(`\nError: ${(error as Error).message}`));
      console.error(chalk.gray('Use --verbose flag for more details'));
    }
    
    process.exit(1);
  }
}

// Handle uncaught errors
process.on('unhandledRejection', (error) => {
  console.error(chalk.red('\nUnhandled error:'));
  console.error(error);
  process.exit(1);
});

// Run the CLI
if (require.main === module) {
  main().catch((error) => {
    console.error(chalk.red('Fatal error:'));
    console.error(error);
    process.exit(1);
  });
}