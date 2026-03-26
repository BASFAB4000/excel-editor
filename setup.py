from setuptools import setup, find_packages

setup(
    name="excel-editor",
    version="0.1.0",
    description="RISE Planungsexcel Editor – liest und bearbeitet Excel-Dateien mit Erhalt der Formatierung.",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    install_requires=[
        "openpyxl>=3.1",
        "pandas>=2.0",
        "pydantic>=2.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-editor=excel_editor.cli:main",
        ],
    },
    python_requires=">=3.8",
)