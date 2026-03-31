import sys
print(f"Python: {sys.executable}")
print(f"Version: {sys.version}")
try:
    import openpyxl
    print(f"openpyxl: {openpyxl.__version__}")
except ImportError:
    print("openpyxl: NOT INSTALLED")
try:
    import pytest
    print(f"pytest: {pytest.__version__}")
except ImportError:
    print("pytest: NOT INSTALLED")

# Write to file since stdout may not be captured
with open("env_check.txt", "w") as f:
    f.write(f"Python: {sys.executable}\n")
    f.write(f"Version: {sys.version}\n")
    try:
        import openpyxl
        f.write(f"openpyxl: {openpyxl.__version__}\n")
    except ImportError:
        f.write("openpyxl: NOT INSTALLED\n")
    try:
        import pytest
        f.write(f"pytest: {pytest.__version__}\n")
    except ImportError:
        f.write("pytest: NOT INSTALLED\n")
