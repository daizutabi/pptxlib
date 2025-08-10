# Troubleshooting

Quick reference for common issues when building, running, or writing docs.

## PowerPoint doesn’t start / COM error

- Confirm PowerPoint is installed
- Quit any already-running PowerPoint instances
- Check if running as administrator affects behavior

## Encoding or font problems

- Ensure the font is installed on the machine
- Specify a fallback font (e.g., `Font(name=...)`)

## Figures don’t appear with mkdocs-nbsync

- Ensure the figure is saved in the notebook (executed and saved)
- Check the Markdown reference: `![alt](../notebooks/example.ipynb){#figure-id}`
- Inspect `mkdocs serve` output for errors

## MkDocs build fails

- Run `uv sync --all-groups` to update dependencies
- Check mkdocs availability with `mkdocs --version`
