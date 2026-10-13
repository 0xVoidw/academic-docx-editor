"""Simple monochrome figures from caller-provided data or graph relations."""

from __future__ import annotations

import shutil
from pathlib import Path

from .models import EditorError, MissingDependency


def bar_chart(data: dict, output: Path) -> dict:
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise MissingDependency("Matplotlib is required: pip install 'academic-docx-editor[figures]'") from exc
    labels = data.get("labels")
    values = data.get("values")
    if not isinstance(labels, list) or not isinstance(values, list) or not labels or len(labels) != len(values):
        raise EditorError("Figure data needs equal-length nonempty labels and values")
    if not all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values):
        raise EditorError("Figure values must be numeric")
    output = Path(output)
    if output.exists() or output.suffix.lower() not in (".png", ".svg", ".pdf"):
        raise EditorError("Choose a new .png, .svg, or .pdf output")
    fig, ax = plt.subplots(figsize=(6.5, 4))
    ax.bar([str(x) for x in labels], values, color="0.72", edgecolor="black", linewidth=0.8)
    ax.set_ylabel(str(data.get("ylabel", "Value")))
    if data.get("title"):
        ax.set_title(str(data["title"]))
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color="0.85", linewidth=0.5)
    ax.set_axisbelow(True)
    fig.tight_layout()
    output.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output, dpi=300, metadata={"Creator": "academic-docx-editor"})
    plt.close(fig)
    return {"output": str(output.resolve()), "format": output.suffix.lower()[1:], "data_points": len(values)}


def flowchart(spec: dict, output: Path) -> dict:
    try:
        import graphviz
    except ImportError as exc:
        raise MissingDependency("graphviz Python package is required: pip install 'academic-docx-editor[figures]'") from exc
    if not shutil.which("dot"):
        raise MissingDependency("Graphviz 'dot' executable is not on PATH")
    nodes, edges = spec.get("nodes"), spec.get("edges")
    if not isinstance(nodes, list) or not isinstance(edges, list) or not nodes:
        raise EditorError("Flowchart spec needs nodes and edges arrays")
    output = Path(output)
    if output.exists() or output.suffix.lower() not in (".png", ".svg", ".pdf"):
        raise EditorError("Choose a new .png, .svg, or .pdf output")
    dot = graphviz.Digraph(format=output.suffix[1:].lower(), graph_attr={"bgcolor": "white", "rankdir": "TB"},
                           node_attr={"shape": "box", "color": "black", "fontcolor": "black", "style": "rounded"},
                           edge_attr={"color": "black"})
    identifiers = set()
    for node in nodes:
        name = node.get("id") if isinstance(node, dict) else None
        if not isinstance(name, str) or not name or name in identifiers:
            raise EditorError("Each node needs a unique nonempty id")
        identifiers.add(name)
        dot.node(name, str(node.get("label", name)))
    for edge in edges:
        if not isinstance(edge, dict) or edge.get("from") not in identifiers or edge.get("to") not in identifiers:
            raise EditorError("Every edge must reference existing node ids")
        dot.edge(edge["from"], edge["to"], label=str(edge.get("label", "")))
    output.parent.mkdir(parents=True, exist_ok=True)
    data = dot.pipe()
    output.write_bytes(data)
    return {"output": str(output.resolve()), "format": output.suffix[1:].lower(),
            "nodes": len(nodes), "edges": len(edges)}
