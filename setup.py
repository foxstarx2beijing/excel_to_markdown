from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excel_to_markdown",
    version="0.1.0",
    author="foxstarx",
    author_email="foxstarx@gmail.com",
    description="Convert Excel/CSV files to Markdown tables",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/excel_to_markdown",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pandas>=2.0.0",
        "openpyxl>=3.0.0",
        "xlrd>=2.0.0",
        "click>=8.0.0",
        "rich>=12.0.0",
        "tabulate>=0.8.0",
    ],
    entry_points={
        "console_scripts": [
            "excel2md=excel_to_md:main",
            "excel2md-gui=gui:main",
        ],
    },
) 