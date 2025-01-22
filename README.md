# TimeStampParser
TextGrid Interval Extractor Script
==================================

This script extracts labeled intervals from TextGrid files and saves the results in an Excel workbook (.xlsx). 
TextGrid files are widely used in linguistic research for annotating audio files, particularly with Praat.

Features
--------
- Extracts "Start Time," "End Time," and "Label" from annotated intervals.
- Processes multiple .textgrid files in a directory.
- Skips intervals with empty labels for cleaner data.
- Outputs results to an easy-to-read Excel file: output_timestamps.xlsx.

Requirements
------------
1. Python 3.6+
2. openpyxl library (install using `pip install openpyxl`)
3. A directory containing .textgrid files.

How to Use
----------
1. Place your .textgrid files in a directory (e.g., `D:\Praat`).
2. Update the `input_directory` variable in the script to point to your directory.
3. Run the script:
   ```python
   python script_name.py
   ```
4. The results will be saved in `output_timestamps.xlsx` in the same directory.

Output Format
-------------
The Excel file will include the following columns:
- File Name: Name of the .textgrid file.
- Start Time: Start time of the labeled interval.
- End Time: End time of the labeled interval.
- Label: The text annotation of the interval.

Example:
| File Name       | Start Time | End Time | Label        |
|------------------|------------|----------|--------------|
| example1.textgrid | 0.00       | 1.25     | Intro        |
| example2.textgrid | 1.50       | 2.75     | Dialogue A   |

Customization
-------------
- To include intervals with empty labels, modify the following line in the script:
  ```python
  if label.strip():
  ```
- To change the output file name or location, update the `output_file` variable.

Known Limitations
-----------------
- Assumes .textgrid files are encoded in UTF-8.
- Skips intervals with empty labels (by default).
- Does not support nested tiers or non-standard TextGrid formats.

Authors
-------
1. Wellington Mendes (UFU, 2025)
   - Email: wellington.mendes@ufu.br
   - Institutional Profile: http://www.portal.ileel.ufu.br/pessoas/docentes/wellington-araujo-mendes-junior
   - Google Scholar: https://scholar.google.com/citations?user=eI4709wAAAAJ&hl=pt-BR

2. Guilherme Felisbino (UFU, 2025)
   - Email: guihperez6@gmail.com

License
-------
This script is distributed under the MIT License. Feel free to modify and use it for your projects.

Contact
-------
For questions, suggestions, or bug reports, contact the authors using the emails above.
