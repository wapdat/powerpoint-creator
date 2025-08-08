---
title: Simple Markdown Presentation
author: John Doe
company: Acme Corp
date: 2025-01-15
---

# Simple Markdown Presentation

## Introduction

Welcome to this presentation created from markdown!

- Markdown is simple to write
- Automatically converts to PowerPoint
- Professional formatting applied

## Key Features

### Automatic Slide Generation

PowerPoint Creator automatically converts your markdown headings into slides:

- H1 becomes title slides
- H2 becomes section dividers
- H3 starts new content slides

### Rich Content Support

You can include various types of content:

1. Ordered lists
2. Unordered lists
3. **Bold text**
4. *Italic text*
5. Code blocks
6. Tables

## Data Visualization

### Sales Data

| Quarter | Revenue | Growth |
|---------|---------|--------|
| Q1 2024 | $1.2M   | 15%    |
| Q2 2024 | $1.5M   | 25%    |
| Q3 2024 | $1.8M   | 20%    |
| Q4 2024 | $2.1M   | 17%    |

## Chart Example

```csv
Month,Sales,Profit
Jan,45,15
Feb,52,18
Mar,58,22
Apr,65,25
May,72,28
Jun,80,32
```

## Code Examples

### Python Example

```python
def generate_presentation(markdown_file):
    """Convert markdown to PowerPoint"""
    with open(markdown_file, 'r') as f:
        content = f.read()
    
    presentation = convert_to_pptx(content)
    presentation.save('output.pptx')
```

## Images

![Company Logo](https://via.placeholder.com/400x200)

## Summary

- Markdown makes presentation creation easy
- Automatic formatting and layout
- Support for charts, tables, and code
- Professional results every time

---

## Thank You

Questions and Discussion

<!-- notes: Remember to emphasize the simplicity of markdown -->