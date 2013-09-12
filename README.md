reqanalysis
===========

Analysis tools for OOI Requirements. It parsed an Excel spreadsheet with CI requirements and
produces an Excel spreadsheet with an analysis report.

## Usage

python reqanalysis.py < in_filename > < out_filename >

The default < in_filename > is defined in the code.

The default < out_filename > is output/reqanalysis_< timestamp >

## Prerequisites

```mkvirtualenv --no-site-packages --python=python2.7 req
easy_install pip
pip install xlrd
pip install xtwt
```
