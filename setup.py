from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="precisiondoc",
    version="0.1.0",
    author="Kay Chiao",
    author_email="kayjiao@163.com",
    description="Document processing and evidence extraction package for precision oncology",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/kaychiao/precisiondoc",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    install_requires=[
        "openai>=0.27.0",
        "python-dotenv>=0.19.0",
        "pandas>=1.3.0",
        "numpy>=1.20.0",
        "pymupdf>=1.19.0",
        "python-docx>=0.8.11",
        "openpyxl>=3.0.7",
        "pillow>=8.2.0",
    ],
    entry_points={
        "console_scripts": [
            "precisiondoc=precisiondoc.main:main",
        ],
    },
    include_package_data=True,
)
