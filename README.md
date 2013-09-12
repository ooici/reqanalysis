reqanalysis
===========

Analysis tools for OOI Requirements. It parsed an Excel spreadsheet with CI requirements and
produces an Excel spreadsheet with an analysis report.

### Usage

python reqanalysis.py {in_filename} {out_filename}

* The default in_filename is defined in the code.
* The default out_filename is output/reqanalysis_{timestamp}

### Prerequisites

    mkvirtualenv --no-site-packages --python=python2.7 req
    easy_install pip
    pip install xlrd
    pip install xtwt
    mkdir output

### Analysis Remarks

* Verification status is taken from the "Group" column of the L4 tab
* A requirement is considered "addressed" if it is R1 or R2 verified or expected in the 
  STC R3, or if one or more "addressed" child requirements exist.
* There are duplicate child requirements in the L2-L3 and L3-L4 tabs. The "Num Parents" column shows the number of parents
* If no link to a parent exists, an L4 requirement is not shown
* Uplinks to L3 interface requirements are not considered
