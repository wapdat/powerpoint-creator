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
  .scriptName('powerpoint-creator')
  .usage('$0 [options]')
  .command('$0', 'Generate professional PowerPoint presentations from JSON data')
  .option('input', {
    alias: 'i',
    type: 'string',
    description: 'Path to JSON input file (or pipe from STDIN)',
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
    description: 'Path to existing PPTX template file (preserves branding)',
  })
  .option('pdf', {
    alias: 'p',
    type: 'boolean',
    description: 'Also generate a PDF version (requires LibreOffice)',
    default: false,
  })
  .option('verbose', {
    alias: 'v',
    type: 'boolean',
    description: 'Show detailed progress and debug information',
    default: false,
  })
  .help()
  .alias('help', 'h')
  .version()
  .alias('version', 'V')
  .example([
    ['$0 --input slides.json --output presentation.pptx', 'Basic usage: Generate from JSON file'],
    ['$0 -i data.json -o report.pptx -v', 'Short flags with verbose output'],
    ['', ''],
    ['# Using templates (preserve corporate branding)', ''],
    ['$0 --input slides.json --template company-template.pptx --output branded.pptx', 'Apply your company template'],
    ['$0 -i q4-data.json -t templates/corporate.pptx -o q4-report.pptx', 'Generate quarterly report with template'],
    ['', ''],
    ['# PDF generation', ''],
    ['$0 --input slides.json --output deck.pptx --pdf', 'Generate both PPTX and PDF'],
    ['$0 -i data.json -o presentation.pptx -p -v', 'PDF with verbose output'],
    ['', ''],
    ['# Input from STDIN (pipe from other commands)', ''],
    ['cat slides.json | $0 --output presentation.pptx', 'Pipe JSON from file'],
    ['curl https://api.example.com/data | $0 -o report.pptx', 'Generate from API response'],
    ['echo \'{"slides":[...]}\' | $0 -o quick.pptx', 'Inline JSON data'],
    ['', ''],
    ['# Batch processing', ''],
    ['for file in *.json; do', '  Multiple presentations'],
    ['  $0 -i "$file" -o "${file%.json}.pptx"', ''],
    ['done', ''],
    ['', ''],
    ['# Advanced examples', ''],
    ['$0 -i <(jq \'.data\' api-response.json) -o filtered.pptx', 'Process JSON with jq first'],
    ['$0 -i report.json -t brand.pptx -o final.pptx --pdf', 'Full pipeline: template + PDF'],
  ])
  .epilogue(`
${chalk.bold('📊 JSON Input Format:')}
  Your JSON file should follow this structure:
  {
    "title": "Presentation Title",
    "author": "Your Name",
    "slides": [
      {
        "layout": "title",
        "title": "Welcome",
        "subtitle": "Subtitle text"
      },
      {
        "layout": "text",
        "title": "Agenda",
        "bullets": ["Point 1", "Point 2", "Point 3"]
      },
      {
        "layout": "chart",
        "title": "Sales Data",
        "chartType": "bar",
        "data": {
          "labels": ["Q1", "Q2", "Q3", "Q4"],
          "datasets": [{
            "label": "Revenue",
            "data": [100, 150, 200, 250]
          }]
        }
      }
    ]
  }

${chalk.bold('🎨 Supported Slide Layouts:')}
  • title    - Title slide with subtitle
  • text     - Bullet points and text
  • image    - Image with caption
  • chart    - Data visualizations (bar, line, pie, area, scatter, doughnut, radar)
  • table    - Structured data tables
  • notes    - Speaker notes slide
  • custom   - Custom layout with positioned elements

${chalk.bold('📋 Tips:')}
  • Use templates to maintain brand consistency
  • HTML tags work in text: <strong>, <em>, <u>
  • Charts support multiple datasets
  • Tables can have custom styling
  • Images can be local paths or URLs

${chalk.bold('🔗 More Information:')}
  Documentation: https://github.com/wapdat/powerpoint-creator
  Examples: https://github.com/wapdat/powerpoint-creator/tree/main/examples
  NPM: https://www.npmjs.com/package/powerpoint-creator
  `)
  .wrap(100)
  .parseSync() as CLIOptions;

/**
 * Display welcome banner
 */
function showBanner(): void {
  if (!argv.verbose) return;
  
  console.log(chalk.blue(`
╔═══════════════════════════════════════════════════════╗
║                                                       ║
║     ${chalk.bold.white('powerpoint-creator')} - PowerPoint Generator      ║
║     ${chalk.gray('Transform JSON → Professional Presentations')}      ║
║                                                       ║
╚═══════════════════════════════════════════════════════╝
  `));
}

/**
 * Main CLI execution function
 */
async function main(): Promise<void> {
  showBanner();
  
  const spinner = ora({
    spinner: 'dots',
    color: 'blue',
  });
  
  try {
    // Handle input source
    let inputData: Presentation;
    
    if (argv.input) {
      // Read from file
      spinner.start(chalk.blue('📂 Reading input file...'));
      
      if (!fs.existsSync(argv.input)) {
        spinner.fail(chalk.red(`❌ Input file not found: ${argv.input}`));
        console.log(chalk.gray('\nTip: Check the file path or use --help for examples'));
        process.exit(1);
      }
      
      const inputContent = fs.readFileSync(argv.input, 'utf-8');
      
      try {
        inputData = JSON.parse(inputContent);
        spinner.succeed(chalk.green(`✅ Input file parsed successfully (${inputData.slides?.length || 0} slides)`));
      } catch (error) {
        spinner.fail(chalk.red('❌ Invalid JSON in input file'));
        if (argv.verbose) {
          console.error(chalk.red('\nError details:'));
          console.error(error);
          console.log(chalk.gray('\nTip: Validate your JSON at https://jsonlint.com'));
        }
        process.exit(1);
      }
    } else {
      // Read from STDIN
      spinner.start(chalk.blue('📥 Reading from STDIN...'));
      
      const chunks: Buffer[] = [];
      
      await new Promise<void>((resolve, reject) => {
        // Don't set encoding to get Buffer chunks
        
        process.stdin.on('data', (chunk: Buffer) => {
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
        inputData = JSON.parse(Buffer.concat(chunks).toString('utf-8'));
        spinner.succeed(chalk.green(`✅ STDIN input parsed successfully (${inputData.slides?.length || 0} slides)`));
      } catch (error) {
        spinner.fail(chalk.red('❌ Invalid JSON from STDIN'));
        if (argv.verbose) {
          console.error(chalk.red('\nError details:'));
          console.error(error);
          console.log(chalk.gray('\nTip: Ensure your piped data is valid JSON'));
        }
        process.exit(1);
      }
    }
    
    // Validate input structure
    if (!inputData.slides || !Array.isArray(inputData.slides)) {
      spinner.fail(chalk.red('❌ Invalid input: missing "slides" array'));
      console.log(chalk.gray('\nExpected structure:'));
      console.log(chalk.gray(JSON.stringify({
        title: "Presentation Title",
        slides: [{ layout: "title", title: "..." }]
      }, null, 2)));
      process.exit(1);
    }
    
    // Validate template if provided
    if (argv.template) {
      spinner.start(chalk.blue('🎨 Validating template file...'));
      
      if (!fs.existsSync(argv.template)) {
        spinner.fail(chalk.red(`❌ Template file not found: ${argv.template}`));
        console.log(chalk.gray('\nTip: Ensure the template path is correct'));
        process.exit(1);
      }
      
      if (!argv.template.endsWith('.pptx')) {
        spinner.fail(chalk.red('❌ Template must be a .pptx file'));
        console.log(chalk.gray('\nTip: Use an existing PowerPoint file as template'));
        process.exit(1);
      }
      
      spinner.succeed(chalk.green('✅ Template file validated'));
    }
    
    // Ensure output directory exists
    const outputDir = path.dirname(argv.output);
    if (!fs.existsSync(outputDir)) {
      if (argv.verbose) {
        console.log(chalk.gray(`Creating output directory: ${outputDir}`));
      }
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    // Show generation details if verbose
    if (argv.verbose) {
      console.log(chalk.gray('\n📋 Generation Details:'));
      console.log(chalk.gray(`  • Title: ${inputData.title || 'Untitled'}`));
      console.log(chalk.gray(`  • Author: ${inputData.author || 'Unknown'}`));
      console.log(chalk.gray(`  • Slides: ${inputData.slides.length}`));
      console.log(chalk.gray(`  • Template: ${argv.template || 'None (default styling)'}`));
      console.log(chalk.gray(`  • Output: ${path.resolve(argv.output)}`));
      
      // Show slide breakdown
      const layoutCounts: Record<string, number> = {};
      inputData.slides.forEach(slide => {
        layoutCounts[slide.layout] = (layoutCounts[slide.layout] || 0) + 1;
      });
      console.log(chalk.gray(`  • Slide types: ${Object.entries(layoutCounts).map(([k, v]) => `${k}(${v})`).join(', ')}`));
    }
    
    // Generate presentation
    spinner.start(chalk.blue('🚀 Generating presentation...'));
    
    const startTime = Date.now();
    
    await generatePresentation({
      inputData,
      outputPath: argv.output,
      templatePath: argv.template,
      convertToPdf: argv.pdf,
      validation: true,
    });
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    
    spinner.succeed(chalk.green(`✅ Presentation generated in ${duration}s: ${chalk.bold(argv.output)}`));
    
    // Show file size
    const stats = fs.statSync(argv.output);
    const fileSize = (stats.size / 1024).toFixed(1);
    console.log(chalk.gray(`   File size: ${fileSize} KB`));
    
    // Handle PDF conversion if requested
    if (argv.pdf) {
      spinner.start(chalk.blue('📄 Converting to PDF...'));
      
      const pdfPath = argv.output.replace(/\.pptx$/, '.pdf');
      
      try {
        // Check if PDF was created
        if (fs.existsSync(pdfPath)) {
          const pdfStats = fs.statSync(pdfPath);
          const pdfSize = (pdfStats.size / 1024).toFixed(1);
          spinner.succeed(chalk.green(`✅ PDF generated: ${chalk.bold(pdfPath)}`));
          console.log(chalk.gray(`   PDF size: ${pdfSize} KB`));
        } else {
          spinner.info(chalk.yellow(`⚠️  PDF conversion requires LibreOffice to be installed`));
          console.log(chalk.gray('\nTo install LibreOffice:'));
          console.log(chalk.gray('  • macOS: brew install --cask libreoffice'));
          console.log(chalk.gray('  • Ubuntu: sudo apt-get install libreoffice'));
          console.log(chalk.gray('  • Windows: Download from libreoffice.org'));
        }
      } catch (error) {
        spinner.warn(chalk.yellow('⚠️  PDF conversion failed'));
        if (argv.verbose) {
          console.error(error);
        }
      }
    }
    
    // Success message
    console.log(chalk.green.bold('\n✨ Success! Your presentation is ready.'));
    
    if (argv.verbose) {
      console.log(chalk.gray('\n📊 Summary:'));
      console.log(chalk.gray(`  • Total slides: ${inputData.slides.length}`));
      console.log(chalk.gray(`  • Generation time: ${duration}s`));
      console.log(chalk.gray(`  • Output location: ${path.resolve(argv.output)}`));
      
      if (argv.template) {
        console.log(chalk.gray(`  • Template applied: ${path.basename(argv.template)}`));
      }
      
      if (argv.pdf) {
        console.log(chalk.gray(`  • PDF location: ${path.resolve(argv.output.replace(/\.pptx$/, '.pdf'))}`));
      }
      
      console.log(chalk.gray('\n💡 Next steps:'));
      console.log(chalk.gray('  • Open the PPTX file in PowerPoint or Google Slides'));
      console.log(chalk.gray('  • Review and customize as needed'));
      console.log(chalk.gray('  • Share or present your creation!'));
    }
    
  } catch (error) {
    spinner.fail(chalk.red('❌ Failed to generate presentation'));
    
    if (argv.verbose) {
      console.error(chalk.red('\n🔍 Error details:'));
      console.error(error);
      console.error(chalk.red('\n📚 Stack trace:'));
      console.error((error as Error).stack);
    } else {
      console.error(chalk.red(`\n❌ Error: ${(error as Error).message}`));
      console.error(chalk.gray('Use --verbose flag for more details'));
    }
    
    // Provide helpful suggestions based on error
    const errorMsg = (error as Error).message.toLowerCase();
    if (errorMsg.includes('validation')) {
      console.log(chalk.yellow('\n💡 Tip: Check your JSON structure matches the expected format'));
      console.log(chalk.gray('   Run with --help to see the correct format'));
    } else if (errorMsg.includes('template')) {
      console.log(chalk.yellow('\n💡 Tip: Ensure your template is a valid .pptx file'));
    } else if (errorMsg.includes('permission')) {
      console.log(chalk.yellow('\n💡 Tip: Check you have write permissions for the output directory'));
    }
    
    process.exit(1);
  }
}

// Handle uncaught errors
process.on('unhandledRejection', (error) => {
  console.error(chalk.red('\n⚠️  Unexpected error:'));
  console.error(error);
  if (!argv.verbose) {
    console.error(chalk.gray('\nRun with --verbose for more details'));
  }
  process.exit(1);
});

// Handle SIGINT (Ctrl+C)
process.on('SIGINT', () => {
  console.log(chalk.yellow('\n\n👋 Generation cancelled by user'));
  process.exit(0);
});

// Run the CLI
if (require.main === module) {
  main().catch((error) => {
    console.error(chalk.red('💥 Fatal error:'));
    console.error(error);
    process.exit(1);
  });
}