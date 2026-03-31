"""Run pytest and write output to file."""
import subprocess
import sys

result = subprocess.run(
    [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"],
    cwd=r"c:\Users\vatsal.gaur\Desktop\SpreadsheetLLm",
    capture_output=True,
    text=True,
    timeout=300,
)
with open("test_results.txt", "w") as f:
    f.write("=== STDOUT ===\n")
    f.write(result.stdout or "(empty)\n")
    f.write("\n=== STDERR ===\n")
    f.write(result.stderr or "(empty)\n")
    f.write(f"\nReturn code: {result.returncode}\n")
