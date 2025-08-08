"""
Setup script for JSON Data Converter
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="array-mate",
    version="1.0.0",
    author="Michael Dehne",
    author_email="",  # Add your email if desired
    description="A GUI application that converts JSON arrays to multiple formats (Excel, CSV, JSON)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/micdehne/array-mate",
    py_modules=["array_mate"],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Utilities",
    ],
    python_requires=">=3.7",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "array-mate=array_mate:main",
        ],
    },
    keywords="json, excel, csv, converter, gui, tkinter, pandas",
    project_urls={
        "Bug Reports": "https://github.com/micdehne/array-mate/issues",
        "Source": "https://github.com/micdehne/array-mate",
    },
)
