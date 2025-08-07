# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**pptx-auto-gen** is a professional TypeScript NPM package for generating PowerPoint presentations programmatically. It creates business-ready PPTX files from JSON input, with support for templates and PDF export.

### Key Features
- Generates professional presentations from JSON
- Supports template-based generation using existing PPTX files
- Multiple slide layouts (title, text, image, chart, table, notes, custom)
- PDF conversion capability
- CLI and programmatic API
- Comprehensive input validation

## Architecture

### Core Modules

1. **src/cli.ts** - Command-line interface using yargs
2. **src/index.ts** - Main API and orchestration
3. **src/renderer.ts** - Slide rendering using PptxGenJS (from scratch)
4. **src/template.ts** - Template processing using pptx-automizer
5. **src/validator.ts** - JSON schema validation using Ajv
6. **src/pdf-converter.ts** - PDF conversion using LibreOffice
7. **src/types.ts** - TypeScript type definitions

### Technology Stack
- **PptxGenJS** - Core presentation generation library
- **pptx-automizer** - Template manipulation and placeholder replacement
- **TypeScript** - Type-safe development
- **Ajv** - JSON schema validation
- **Yargs** - CLI argument parsing

## Common Development Tasks

### Build the project
```bash
npm run build
```

### Run in development mode
```bash
npm run dev
```

### Test the CLI locally
```bash
npm run build
node dist/cli.js --input examples/slides.json --output test.pptx
```

### Test programmatic API
```bash
npm run build
node -e "const { generatePresentation } = require('./dist'); generatePresentation({ inputPath: 'examples/simple.json', outputPath: 'test.pptx' }).then(() => console.log('Done'));"
```

### Generate example presentation
```bash
npm run example
```

## Key Implementation Details

### Slide Rendering Strategy

The package uses two different approaches based on whether a template is provided:

1. **Without Template** (src/renderer.ts):
   - Uses PptxGenJS to create presentations from scratch
   - Applies professional color schemes and styling
   - Full control over layout and design

2. **With Template** (src/template.ts):
   - Uses pptx-automizer to manipulate existing PPTX files
   - Preserves template styling and master slides
   - Replaces placeholders with content

### Input Validation

All input is validated against a comprehensive JSON schema before processing:
- Required fields validation
- Type checking for all properties
- Layout-specific field requirements
- Color format validation
- Data consistency checks (e.g., chart data length matching labels)

### PDF Conversion

PDF conversion is handled through multiple fallback methods:
1. LibreOffice (primary method, cross-platform)
2. Native PowerPoint (Windows/Mac, if available)
3. Puppeteer (future implementation)

### Error Handling

The package implements comprehensive error handling:
- Detailed validation error messages with field paths
- Graceful fallbacks for missing resources (images, templates)
- Informative CLI output with verbose mode
- Proper cleanup of temporary files

## Adding New Features

### Adding a New Slide Layout

1. Add the layout type to `SlideLayout` in src/types.ts
2. Create the interface for the new slide type in src/types.ts
3. Add rendering logic in src/renderer.ts `renderSlide()` method
4. Add template processing in src/template.ts `processSlide()` method
5. Add validation rules in src/validator.ts `validateSlide()` method
6. Update the README with the new layout documentation

### Adding a New Chart Type

1. Add to `ChartType` enum in src/types.ts
2. Update chart type mapping in src/renderer.ts `renderChartSlide()`
3. Add validation in src/validator.ts `isValidChartType()`
4. Update documentation

## Testing Considerations

When testing modifications:
1. Test both with and without templates
2. Validate JSON input with intentional errors
3. Test PDF conversion on different platforms
4. Verify CLI and API interfaces work correctly
5. Check error messages are helpful and accurate

## Performance Optimization

For large presentations:
- Images are processed asynchronously
- Template processing uses efficient placeholder replacement
- Validation is done upfront to fail fast
- Temporary files are cleaned up immediately

## Common Issues and Solutions

### Issue: Template placeholders not being replaced
- Ensure placeholder types match exactly (title, subtitle, body, etc.)
- Check that the template has the expected placeholder structure

### Issue: Charts not rendering correctly
- Verify data array lengths match label counts
- Ensure chart type is supported
- Check color formats are valid

### Issue: PDF conversion failing
- Verify LibreOffice is installed and in PATH
- Check file permissions for output directory
- Ensure input PPTX is valid

## Code Style Guidelines

- Use TypeScript strict mode
- Implement comprehensive error handling
- Add JSDoc comments for public APIs
- Follow existing patterns for consistency
- Validate all external input
- Clean up resources (files, buffers) properly

## Future Enhancements

Potential areas for expansion:
- Animation and transition support
- Advanced chart customization
- Video/audio embedding
- Cloud storage integration
- Real-time collaboration features
- Additional export formats