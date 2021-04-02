# Word2CSV Converter

Copy following items into a local folder:

    - file:         word2csv.bat
    - directory:   word2csv

In your local folder, drag & drop one or more ".docx" files onto the "word2csv.bat" file.

A command window opens and ask to confirm the output path for the converted CSV files (the chosen path will be remembered)
The program will show a progress bar for each file while it is converted.

## Developer Notes

The main program is written in Python and located in `word2csv\src`

The `word2csv.bat` expects a [WinPython](https://winpython.github.io/) installation located at `word2csv\Python`

The following Python modules must be installed in the WinPython environment:
- lxml.etree
- pandas
- numpy
- configparser
- alive_progress
