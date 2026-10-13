import pytest

from academic_docx_editor.figures import bar_chart, flowchart
from academic_docx_editor.models import MissingDependency


@pytest.mark.parametrize('extension', ['png', 'svg', 'pdf'])
def test_monochrome_bar_chart(tmp_path, extension):
    pytest.importorskip('matplotlib')
    output = tmp_path / f'plot.{extension}'
    result = bar_chart({'labels': ['A', 'B'], 'values': [2, 4], 'ylabel': 'Count'}, output)
    assert result['data_points'] == 2
    assert output.stat().st_size > 100


def test_graphviz_missing_dot_reported(tmp_path, monkeypatch):
    import academic_docx_editor.figures as figures
    if not __import__('importlib').util.find_spec('graphviz'):
        with pytest.raises(MissingDependency, match='graphviz Python'):
            flowchart({'nodes': [{'id': 'A'}], 'edges': []}, tmp_path / 'flow.svg')
