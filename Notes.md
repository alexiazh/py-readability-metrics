# Excel Readability Scorer

## Table of Contents

1. [Description](#desc)
2. [Installation](#install)
    * [Virtual environment](#venv)
    * [Libraries & dependencies](#dep)
3. [Usage](#usage)
4. [References](#ref)

## Description <a name="desc"></a>

The script scores the *readability* of text in Excel files.

The readability metrics used are:
- Flesch Reading Ease (FRES)
- Flesch Kincaid Grade Level (FKGL)
- Gunning Fog Index (GF)
- SMOG
- Coleman Liau Index (CLI)
- Automated Readability Index (ARI)

## Installation <a name="install"></a>

### Virtual environment <a name="venv"></a>

When using pip it is generally recommended to install packages in a virtual environment to avoid modifying system state:

```bash
$ conda create -n readability
$ conda activate readability
```

### Libraries & dependencies <a name="dep"></a>

```bash
$ pip install py-readability-metrics
```

Run the Python interpreter and type the commands:

```python
>>> import nltk
>>> nltk.download()
```

## Usage <a name="usage"></a>

```bash
$ python excel_scorer.py
```

## References <a name="ref"></a>

- [1] https://pypi.org/project/py-readability-metrics/
- [2] https://github.com/cdimascio/py-readability-metrics