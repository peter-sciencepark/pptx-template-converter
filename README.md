# pptx-template-converter

Converts PowerPoint presentations from an old template/layout to a new one.

Built for migrating Science Park presentations to the new 2023 brand template.

## How it works

1. Opens the source `.pptx` and the new `.potx` template
2. Classifies each slide (title, chapter, content, quote, closing, etc.)
3. Maps each slide to the best matching layout in the new template
4. Transfers text content, bullet lists, and images
5. Saves a new `.pptx` with the new branding

## Usage

```bash
pip install -r requirements.txt

python3 convert.py presentation.pptx --template "Ny mall.potx" -o output.pptx
```

## Slide type mapping

| Detected type | New template layout |
|---|---|
| Title / cover | `1 - Rubrikbild logo` |
| Chapter heading | `1 - Kapitelrubrik` |
| Quote | `11 - Midicitat blå` |
| Content (text) | `4 - Innehåll blank` |
| Content (with image) | `4 - Bild höger` |
| Closing | `14 - Slogan hav` |

## Limitations

- Complex diagrams and SmartArt are not automatically converted
- Tables may need manual adjustment
- Animations and transitions are not preserved
- Some slides may need manual touch-up after conversion
