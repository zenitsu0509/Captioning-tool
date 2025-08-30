# Captioning Helper Tool

This lightweight tool helps you generate structured captions for your images and save them directly into `captions_prompts.xlsx` in row order. It uses your `Caption_guide_reg.xlsx` to populate dropdowns for fast, consistent entry.

## What it does

- Scans an images folder (default: `../1_indian_woman`) and sorts files by number.
- Loads caption options from `../Caption_guide_reg.xlsx`.
- Lets you pick: Body Portion, Race, Pose, Dress Types, Location, Props (optional).
- Writes a single caption string per image to `../captions_prompts.xlsx` (row 1 = image 1, row 2 = image 2, ...). Preserves existing captions.
- Optionally marks images with body artifacts in a separate sheet named `artifacts`.

Also included: Automatic mode (Option A) using CLIP to auto-fill fields from your guide and write captions to Excel in order. See below.

## Quick start

1) Install dependencies (in a Python 3.9+ environment):

   - pandas
   - openpyxl
   - pillow

2) Run the app from this folder:

   python app.py

3) Controls:

   - Next: Right Arrow or Enter
   - Previous: Left Arrow
   - Save: Ctrl+S (also happens on Save & Next)
   - Save & Next: Button or Enter
   - Toggle Artifact flag: Alt+A

### Automatic mode (Option A)

Install additional packages (CPU works; GPU accelerates):

   pip install -r requirements-clip.txt

Run on your folder, writing to Excel in row order:

   python auto_caption_clip.py --images_dir "../1_indian_woman" --guide "../Caption_guide_reg.xlsx" --output "../captions_prompts.xlsx" --artifact

Options:

- --limit N: process first N images
- --props_max K: max props selected (default 5)
- --props_floor 0.24 and --props_band 0.02: thresholds for props multi-select
- --model ViT-B-32 --pretrained openai: change CLIP backbone

## Assumptions

- `Caption_guide_reg.xlsx` has columns named: Body Portion, Race, Pose, Dress Types, Location, Props. If arranged differently, the tool tries to infer options; you can still type new values into the dropdowns.
- `captions_prompts.xlsx` has no header row; row 1 corresponds to the first image when sorted numerically.
- Filenames contain a numeric index (e.g., `indian woman_0001.png`). The tool sorts by this number to match row order.

## CLI arguments (optional)

You can override defaults:

- --images_dir "../1_indian_woman"
- --guide "../Caption_guide_reg.xlsx"
- --output "../captions_prompts.xlsx"

## Cleanup

After your approval, you may delete images manually or use your own process. This tool does not delete images automatically.

## Troubleshooting

- If Excel files are open in another program, saving may fail. Close them and try again.
- If dropdowns are empty, check the guide file path and sheet structure.
