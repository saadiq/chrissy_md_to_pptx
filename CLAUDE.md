# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Convert markdown to PowerPoint
uv run md2pptx.py slides.md

# With custom output path
uv run md2pptx.py slides.md -o presentation.pptx

# With template for styling
uv run md2pptx.py slides.md -t template.pptx -o output.pptx
```

## Architecture

Single-file Python script (`md2pptx.py`) using inline script metadata (PEP 723) for dependencies. Runs directly with `uv run` without virtual environment setup.

### Pipeline

1. **Parsing** (`parse_markdown`) - Splits markdown by `---` separators, identifies slide types:
   - Title slides: `## Slide N: Title Slide` with `**Title** *Subtitle*` pattern
   - Section dividers: `# SECTION N: NAME`
   - Content slides: `## Slide N: Title` followed by content

2. **Content extraction** - Helper functions parse tables, bullets, and screenshot placeholders from each slide's content block

3. **Slide generation** - Each slide type has a dedicated `add_*_slide` function that builds the PowerPoint shapes using python-pptx

### Markdown Format

- Slides separated by `---`
- Title slides require `Title Slide` in the header and `**Title** *Subtitle*` pattern
- Section dividers use `# SECTION N: NAME (X slides)` format
- Content slides use `## Slide N: Title` format
- Screenshot placeholders: `📸 **[SCREENSHOT PLACEHOLDER]:** description`

### Color Palette

Constants defined at module level: `NAVY`, `SLATE`, `SILVER`, `OFF_WHITE`, `WHITE`
