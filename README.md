# PowerPoint Creator 🎯

[![npm version](https://badge.fury.io/js/powerpoint-creator.svg)](https://www.npmjs.com/package/powerpoint-creator)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/badge/node-%3E%3D16.0.0-brightgreen)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.3-blue)](https://www.typescriptlang.org/)

> Transform structured JSON data into professional PowerPoint presentations with a single command. Perfect for automation, reporting, and bulk presentation generation.

## 🌟 Key Features

- **📝 Markdown to PowerPoint** - Convert markdown documents directly to presentations
- **📊 7 Slide Layouts** - Title, text, charts, tables, images, notes, and custom layouts
- **🎨 Professional Themes** - Business-ready color schemes and formatting
- **📈 Rich Charts** - Bar, line, pie, area, scatter, doughnut, and radar charts
- **🏢 Template Support** - Apply corporate templates to maintain brand consistency
- **🎯 Smart Content Detection** - Automatically detects charts from CSV/table data
- **📐 Perfect Spacing** - Automatically positions content with professional margins
- **🔄 Batch Processing** - Generate multiple presentations programmatically
- **📄 PDF Export** - Optional PDF generation via LibreOffice
- **✅ No Repair Needed** - Generates clean PPTX files that open without errors

## 📥 Installation

### Global Installation (Recommended)
```bash
npm install -g powerpoint-creator
```

### Local Installation
```bash
npm install powerpoint-creator
```

### Development
```bash
git clone https://github.com/wapdat/powerpoint-creator.git
cd powerpoint-creator
npm install
npm run build
```

## 🚀 Quick Start

### Basic Usage
```bash
powerpoint-creator -i presentation.json -o output.pptx
```

### Markdown to PowerPoint
```bash
powerpoint-creator -m document.md -o presentation.pptx
```

### With Corporate Template
```bash
powerpoint-creator -i data.json -t company-template.pptx -o report.pptx
```

### From STDIN
```bash
cat data.json | powerpoint-creator -o presentation.pptx
```

### Generate PDF
```bash
powerpoint-creator -i slides.json -o deck.pptx --pdf
```

## 📝 Markdown Support

### Overview
PowerPoint Creator can convert markdown documents directly to presentations. It automatically:
- Converts headings to slide titles and sections
- Detects and creates charts from CSV data and tables
- Formats lists, emphasis, and code blocks
- Splits long content across multiple slides
- Processes YAML frontmatter for metadata

### Markdown Structure
```markdown
---
title: Presentation Title
author: Your Name
company: Company Name
date: 2025-01-15
theme: professional
---

# Main Title Slide

## Section Divider

### Content Slide Title

- Bullet point 1
- Bullet point 2
- **Bold text**
- *Italic text*

### Data Table

| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |

### Chart from CSV

\`\`\`csv
Month,Sales,Profit
Jan,100,20
Feb,120,25
Mar,150,35
\`\`\`

![Image Caption](image.png)

<!-- notes: Speaker notes go here -->
```

### Conversion Rules
- **H1 (#)** → Title slides
- **H2 (##)** → Section dividers with colored background
- **H3 (###)** → New content slide or emphasized bullet
- **Lists** → Bullet points with proper indentation
- **Tables** → Table slides or charts (if numeric)
- **Code blocks** → Notes slides or charts (if CSV/JSON)
- **Images** → Image slides with captions
- **Blockquotes** → Notes slides
- **Horizontal rules (---)** → Slide breaks
- **HTML comments** → Speaker notes or directives

## 📋 JSON Structure

### Complete Presentation Example
```json
{
  "title": "Q4 2024 Business Review",
  "author": "John Smith",
  "company": "Acme Corp",
  "slides": [
    {
      "layout": "title",
      "title": "Q4 2024 Results",
      "subtitle": "Record Breaking Quarter",
      "author": "Leadership Team",
      "date": "January 2025"
    },
    {
      "layout": "text",
      "title": "Key Achievements",
      "bullets": [
        "Revenue exceeded targets by 15%",
        "Launched 3 new products successfully",
        "Customer satisfaction increased to 92%"
      ]
    },
    {
      "layout": "chart",
      "title": "Revenue Growth",
      "chartType": "bar",
      "data": {
        "labels": ["Q1", "Q2", "Q3", "Q4"],
        "datasets": [{
          "label": "Revenue ($M)",
          "data": [45, 52, 58, 67]
        }]
      }
    },
    {
      "layout": "table",
      "title": "Regional Performance",
      "headers": ["Region", "Revenue", "Growth"],
      "tableData": [
        ["North America", "$30M", "+20%"],
        ["Europe", "$22M", "+15%"],
        ["Asia Pacific", "$15M", "+35%"]
      ]
    }
  ]
}
```

## 🎨 Slide Types

### 1. Title Slide
Creates opening or section divider slides with customizable backgrounds.

```json
{
  "layout": "title",
  "title": "Main Title",
  "subtitle": "Subtitle Text",
  "author": "Presenter Name",
  "date": "January 2025",
  "backgroundColor": "#2C3E50",
  "notes": "Speaker notes here"
}
```

### 2. Text/Bullet Slide
Perfect for agendas, lists, and text-heavy content.

```json
{
  "layout": "text",
  "title": "Agenda",
  "bullets": [
    "Introduction and Overview",
    "Q4 Performance Metrics",
    "Strategic Initiatives",
    "2025 Roadmap"
  ],
  "level": [0, 0, 1, 1],
  "notes": "Cover each point for 5 minutes"
}
```

### 3. Chart Slide
Visualize data with professional charts.

```json
{
  "layout": "chart",
  "title": "Sales Performance",
  "chartType": "line",
  "data": {
    "labels": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    "datasets": [
      {
        "label": "2024 Sales",
        "data": [30, 45, 60, 70, 85, 95]
      },
      {
        "label": "2023 Sales",
        "data": [25, 35, 45, 55, 65, 75]
      }
    ]
  },
  "options": {
    "showLegend": true,
    "legendPosition": "bottom"
  }
}
```

**Supported Chart Types:**
- `bar` - Vertical bar charts
- `line` - Line graphs
- `pie` - Pie charts
- `area` - Area charts
- `scatter` - Scatter plots
- `doughnut` - Doughnut charts
- `radar` - Radar/spider charts

### 4. Table Slide
Display structured data in professional tables.

```json
{
  "layout": "table",
  "title": "Project Status",
  "headers": ["Project", "Status", "Completion", "Owner"],
  "tableData": [
    ["Website Redesign", "On Track", "75%", "Sarah"],
    ["Mobile App", "At Risk", "45%", "John"],
    ["API Development", "Complete", "100%", "Mike"]
  ],
  "styling": {
    "headerBackground": "#2C3E50",
    "headerTextColor": "#FFFFFF",
    "alternateRows": true
  }
}
```

### 5. Image Slide
Include images with optional captions.

```json
{
  "layout": "image",
  "title": "Product Screenshot",
  "imagePath": "./screenshots/dashboard.png",
  "caption": "New dashboard interface",
  "sizing": "contain",
  "notes": "Highlight the improved UX"
}
```

### 6. Notes Slide
Text-only slides for detailed notes.

```json
{
  "layout": "notes",
  "title": "Implementation Notes",
  "content": "Detailed implementation plan:\n1. Phase 1: Planning\n2. Phase 2: Development\n3. Phase 3: Testing\n4. Phase 4: Deployment"
}
```

### 7. Custom Slide
Advanced layouts with positioned elements.

```json
{
  "layout": "custom",
  "title": "Custom Layout",
  "elements": [
    {
      "type": "text",
      "content": "Custom positioned text",
      "x": 1,
      "y": 2,
      "width": 4,
      "height": 1
    },
    {
      "type": "shape",
      "content": {
        "type": "rect",
        "fill": "#3498DB"
      },
      "x": 6,
      "y": 2,
      "width": 3,
      "height": 3
    }
  ]
}
```

## 🎨 Professional Styling

### Default Color Palette
The package applies a professional business color scheme automatically:
- **Primary**: Dark blue-gray (#2C3E50)
- **Charts**: Professional grayscale palette
- **Tables**: Navy headers with clean alternating rows
- **Backgrounds**: Sophisticated solid colors

### Text Formatting
Use HTML tags in any text field:
- `<strong>Bold text</strong>`
- `<em>Italic text</em>`
- `<u>Underlined text</u>`

### Slide Dimensions
- **Format**: 16:9 widescreen (10" × 7.5")
- **Safe margins**: 0.75" on all sides
- **Title area**: 0.4" - 1.2" from top
- **Content area**: 1.3" - 5.8" from top

## 📂 Example Presentations

We've included two complete example presentations:

1. **[Business Report Example](examples/business-report.json)** - Quarterly business review with charts and tables
   - [View Generated PPTX](examples/business-report-example.pptx)

2. **[Product Launch Example](examples/product-launch.json)** - Product announcement with mixed content types
   - [View Generated PPTX](examples/product-launch-example.pptx)

### Generate Examples
```bash
# Generate business report
powerpoint-creator -i examples/business-report.json -o my-report.pptx

# Generate product launch deck
powerpoint-creator -i examples/product-launch.json -o my-launch.pptx
```

## 🔧 Advanced Usage

### Batch Processing
```bash
# Process multiple JSON files
for file in data/*.json; do
  powerpoint-creator -i "$file" -o "output/$(basename $file .json).pptx"
done
```

### API Integration
```bash
# Generate from API response
curl https://api.example.com/report | powerpoint-creator -o report.pptx
```

### Template Application
```bash
# Apply branding template
powerpoint-creator -i content.json -t templates/corporate.pptx -o branded.pptx
```

### With Data Processing
```bash
# Process with jq first
cat raw-data.json | jq '.presentations[0]' | powerpoint-creator -o filtered.pptx
```

## 🛠️ CLI Options

| Option | Alias | Description | Default |
|--------|-------|-------------|---------|
| `--input` | `-i` | Input JSON file path | stdin |
| `--output` | `-o` | Output PPTX file path | output.pptx |
| `--template` | `-t` | Template PPTX file path | none |
| `--pdf` | `-p` | Also generate PDF | false |
| `--verbose` | `-v` | Show detailed progress | false |
| `--help` | `-h` | Show help message | - |
| `--version` | `-V` | Show version number | - |

## 📦 Programmatic Usage

```javascript
const { PresentationGenerator } = require('powerpoint-creator');

const presentation = {
  title: "My Presentation",
  slides: [
    {
      layout: "title",
      title: "Welcome",
      subtitle: "Let's get started"
    }
  ]
};

const generator = new PresentationGenerator();
const buffer = await generator.generate(presentation);
fs.writeFileSync('output.pptx', buffer);
```

## 🔍 Troubleshooting

### Common Issues

**Issue**: "Command not found" after installation
```bash
# Ensure npm global bin is in PATH
export PATH=$PATH:$(npm prefix -g)/bin
```

**Issue**: PDF generation not working
```bash
# Install LibreOffice
# macOS:
brew install libreoffice
# Ubuntu:
sudo apt-get install libreoffice
```

**Issue**: Images not loading
- Use absolute paths for local images
- Ensure image files exist and are accessible
- For URLs, ensure they're publicly accessible

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) v4.0.1
- Template support via [pptx-automizer](https://github.com/singerla/pptx-automizer)
- CLI powered by [Yargs](https://yargs.js.org/)
- Styled with [Chalk](https://github.com/chalk/chalk)

## 📞 Support

- **Documentation**: [GitHub Wiki](https://github.com/wapdat/powerpoint-creator/wiki)
- **Issues**: [GitHub Issues](https://github.com/wapdat/powerpoint-creator/issues)
- **NPM**: [npmjs.com/package/powerpoint-creator](https://www.npmjs.com/package/powerpoint-creator)

---

**Made with ❤️ for automating presentations**