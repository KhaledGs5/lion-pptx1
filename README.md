# lion-pptx1

Graphic-heavy 5-slide lions presentation generated programmatically.

## Files
- `generate_pptx.py` – reproducible script that creates the presentation
- `lions-presentation.pptx` – generated output file at repo root
- `assets/` – source SVG artwork created for the project

## Regenerate locally
1. Install dependency:
   ```bash
   pip install python-pptx
   ```
2. Generate the presentation:
   ```bash
   python generate_pptx.py
   ```

The script will produce `lions-presentation.pptx` with exactly 5 slides and speaker notes on each slide.

## Notes on visuals and licensing
- All visuals are built from PowerPoint vector shapes and original local SVG assets in `assets/`.
- No external third-party images are embedded.
