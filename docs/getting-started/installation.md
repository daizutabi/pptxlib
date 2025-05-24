# Installation

## Prerequisites

Before installing pptxlib, ensure you have the following:

- Windows operating system
- Microsoft PowerPoint installed on your system
- Python 3.11 or higher

## Installation Methods

### Using uv

The recommended way to install pptxlib is using uv:

```bash
uv pip install pptxlib
```

### Using pip

An alternative way to install pptxlib is using pip:

```bash
pip install pptxlib
```

## Verifying Installation

To verify the installation, you can run Python and import the package:

```python
from pptxlib import App

# Check if PowerPoint is available
if App.is_app_available():
    print("pptxlib is installed and PowerPoint is available")
else:
    print("pptxlib is installed but PowerPoint is not available")
```
