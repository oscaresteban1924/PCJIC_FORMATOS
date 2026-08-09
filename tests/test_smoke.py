import pytest
import subprocess
import sys


def test_smoke_script_execution():
    result = subprocess.run([sys.executable, "scripts/smoke_test.py"], capture_output=True, text=True)
    assert result.returncode == 0
    assert "OK" in result.stdout
