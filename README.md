# python-pptx-merger

A Python library for merging PowerPoint (.pptx) files.

## Installation

1. Install libgdiplus:
   - On MacOS: `brew install mono-libgdiplus`
   - On Ubuntu: `sudo apt-get install libgdiplus`

2. Install the official .NET 8.0 runtime from [Microsoft's website](https://dotnet.microsoft.com/download/dotnet/8.0)

3. Install the package via pip:
   ```bash
   pip install python-pptx-merger
   ```

## Usage

### Merge Whole Presentations

```python
from pptx_merger import Merger

merger = Merger()
merged_doc = merger.merge_slides([src_doc_1, src_doc_2])
```

### Merge Specific Slides

```python
from pptx_merger import SlideRef, Merger

merger = Merger()
slide_refs = [
    SlideRef(doc_idx=0, slide_idx=0),
    SlideRef(doc_idx=0, slide_idx=1),
    SlideRef(doc_idx=1, slide_idx=3),
    SlideRef(doc_idx=0, slide_idx=3),  # Note that slide refs can be out of order
]
merged_doc = merger.merge_slides([src_doc_1, src_doc_2], slide_refs)
```
