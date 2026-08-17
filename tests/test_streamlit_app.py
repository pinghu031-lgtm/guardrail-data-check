from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_streamlit_app_renders_public_upload_interface():
    app_path = Path(__file__).parents[1] / "app.py"

    app = AppTest.from_file(str(app_path)).run(timeout=20)

    assert not app.exception
    assert app.title[0].value == "护栏数据异常检查工具"
    assert len(app.get("file_uploader")) == 1
    assert any(button.label == "开始处理" for button in app.button)
