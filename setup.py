"""Setup script for PD Generator."""

from setuptools import setup, find_packages

setup(
    name="pd-generator",
    version="1.0.0",
    description="CLI tool for generating university project posters",
    author="Your Name",
    python_requires=">=3.8",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    install_requires=[
        "reportlab>=4.0.0",
        "Pillow>=10.0.0",
        "openpyxl>=3.1.0",
        "PyYAML>=6.0",
    ],
    entry_points={
        "console_scripts": [
            "pd-generator=pd_generator.cli:main",
        ],
    },
)
