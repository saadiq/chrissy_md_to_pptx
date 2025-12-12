# md2pptx

Convert markdown slide decks to PowerPoint presentations.

## Requirements

- [uv](https://docs.astral.sh/uv/) (recommended) or Python 3.8+

## Usage

```bash
# Basic usage
uv run md2pptx.py slides.md

# Specify output file
uv run md2pptx.py slides.md -o presentation.pptx

# Use a template for styling
uv run md2pptx.py slides.md -t company-template.pptx -o output.pptx
```

## Markdown Format

Slides are separated by `---` (horizontal rule). The converter recognizes:

### Title Slide
```markdown
## Slide 1: Title Slide

**Presentation Title** *Subtitle text*
```

### Section Divider
```markdown
# SECTION 1: INTRODUCTION (4 slides)
```

### Content Slide
```markdown
## Slide 2: Slide Title

Regular paragraph text.

| Column 1 | Column 2 |
| :------- | :------- |
| Data     | Data     |

- Bullet point
- Another bullet
  - Nested bullet

1. Numbered item
2. Another item
```

### Supported Elements

| Element | Markdown |
|---------|----------|
| Bold | `**text**` |
| Italic | `*text*` |
| Tables | Pipe-delimited tables |
| Bullets | `- item` or `* item` |
| Numbered lists | `1. item` |
| Checkboxes | `- [ ] item` |
| Links | `[text](url)` (text preserved) |
| Image placeholders | `📸 **[SCREENSHOT PLACEHOLDER]:** description` |

### Screenshot Placeholders

Mark where images should go:

```markdown
📸 **[SCREENSHOT PLACEHOLDER]:** Login screen with username field highlighted
```

This renders as a gray placeholder box with the description, making it clear where screenshots need to be added later.

### Ignored Elements

- Document metadata at the start (title, subtitle before first `---`)

## Templates

Pass a `.pptx` template file with `-t` to use custom:
- Slide layouts
- Theme colors and fonts
- Master slide backgrounds/logos

The converter uses blank layouts (layout index 6) by default.

## Example

```bash
uv run md2pptx.py "Annual Review Training Deck - Draft.md"
# Creates: Annual Review Training Deck - Draft.pptx
```
