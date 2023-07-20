# MailSearcher
A simple GUI tool to index or search your local email file (.eml).

## Requirements
Here are the necessary packages if you want to run or compile the program to an executable file.
1. [pytz](https://pythonhosted.org/pytz/)
2. [python-dateutil](https://github.com/dateutil/dateutil)
3. [Beautiful Soup 4](https://www.crummy.com/software/BeautifulSoup/)

You can install the packages via PyPI
```sh
pip3 install pytz && pip3 install python-dateutil && pip3 install beautifulsoup4
```

## Usage
Create a config file (config.txt) first. Then run the mailsearcher.py file or use the pre-compiled executable file

## Configuration
Since the config file use `;` to separate the items, please **DO NOT** use `;` in the email filename. The config file **MUST** be stored in the **SAME** directory as the Python file or the executable file.

### Configuration explanation
|Item|Explanation|Remark|
|--|--|--|
|path|Directory that you store your .eml files|Support multiple subdirectories|
|threading|Perform the search action linearly or concurrently|Since there is GIL in Python so it is not real concurrently|
|timezone|Received time shown in the GUI|Check the time zone via this [wiki](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones), or via `import pytz; print(pytz.all_timezones)`|
|year|Allowed search range in the GUI|`now` is also accepted|

### Sample config file
```
path=C:\Users\user\Desktop\Email;
threading=true;
timezone=Asia/Tokyo;
year=2019-now
```

## Compilation
In case you want to compile the program to the executable file by yourself, you may first install [PyInstaller](https://pyinstaller.org/en/stable/) and then run the following command.
```sh
pyinstaller --onefile --windowed -n "MailSearcher" -i "icon.ico" mailsearcher.py
```
