"""Start the Streamlit app headlessly, as Streamlit Community Cloud would, and check it renders without errors."""

from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parents[1] / "app.py")


def test_app_starts_cleanly():
    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    assert "Monthly attendance marking" in "".join(m.value for m in at.markdown)
