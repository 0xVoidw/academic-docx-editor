# Academic figures

`audit` gives each displayed body image an `image-N` ID, paragraph index, inline/floating status, media part, and size in EMU. `replace-image` accepts a unique inline PNG/JPEG image whose media part is not shared; replacement must use the same file extension. The OOXML drawing markup, dimensions, and neighboring caption remain byte-identical. Floating drawings and embedded charts are protected and require manual handling.

`figure-bar` takes caller-supplied `labels` and numeric `values`; `figure-flowchart` takes caller-supplied `nodes` and `edges`. Both generate monochrome PNG, SVG, or PDF outputs. Do not alter values, relationships, or conclusions while styling. Matplotlib and Graphviz are optional; flowcharts additionally require the Graphviz `dot` executable.

Prefer vector SVG/PDF for standalone publication figures and high-resolution PNG for insertion into Word. The current image replacement command intentionally supports only existing PNG/JPEG media in DOCX. Review accessibility descriptions, numbering, scales, units, and captions manually.
