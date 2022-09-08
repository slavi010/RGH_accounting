# Description
This script is used to process comparisons between lists of values and find pair opposite values on .xlsx files.

Example of process done:

| Original values  | Generated paired number |
|:----------------:|:-----------------------:|
|        5         |            1            |
|        -2        |            2            |
|        -5        |            1            |
|        3         |                         |
|        -5        |                         |
|        -2        |            2            |
|        2         |            2            |
|        2         |            2            |

# Requirements
- Python 3.9+

# Installation

```bash
pip install -r requirements.txt
```

# Quick start

Using the following command, you can run with the pre-defined parameters:
```bash
python cli.py input.xlsx -o output.xlsx -t 'sheet 1'
```

For more information and customize parameters, please see help:
```bash
python cli.py --help
```

# Contributors
- Main author
  - Sviatoslav Besnard
  - Data Analyst Trainee
  - pro@slavi.dev
- Original idea and draft script
  - Hugo Grau
  - hugo.grau@radissonhotels.com