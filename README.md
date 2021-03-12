# PyOpenRPA
This repo contains Python-based robot working on Windows which utilizes PyOpenRPA library for searching and extracting Yandex results in a separate file.

## NOTE! Selenium web driver and Google Chrome portable version were removed from this repo intentionally because the robot was published for reference purposes only.

This robot uses Python 3.7.2 with PyOpenRPA library, which were not commited to Git due to file size limitations. 
In order to test workability, request full archive Project.zip from iduseev and unzip it on your local machine as per instructions below. 
_________________________________________________________________________________________________________________________________________________________________
# Welcome to search_extractor.py!

This is a robot, which takes configuration file, searches for the request on Yandex's first page, extracts search results, takes screenshots, and saves them in .rtf file.

results.rtf is the robot output file and is stored in /builds folder

_________________________________________________________________________________________________________________________________________________________________

## For robot implementation, following technologies and libraries were used:

Selenium for fetching search results and screenshots taking

win32clipboard to paste screenshots

pathlib for paths composition

UIDesktop Module to set focus on WordPad window

Popen subprocess module for opening wordpad.exe

Image module to open, convert and save screenshots in appropriate format

_________________________________________________________________________________________________________________________________________________________________
## Usage:

1. Unzip archive Project.zip;
2. In \sources open settings.py and replace SEARCH_PHRASE by your intended search request;
3. In \builds run run.cmd file;
4. Result will be saved in results.rtf in /builds folder;
5. Everything should work correctly outside of box


_________________________________________________________________________________________________________________________________________________________________
## Documentation used:

UIDesktop
---------------
https://gitlab.com/UnicodeLabs/OpenRPA/-/blob/master/Wiki/05.2.-Theory-&-practice.-Desktop-app-UI-access-(win32-and-UI-automation-dlls).md#UIOSelector_FocusHighlight

https://gitlab.com/UnicodeLabs/OpenRPA/-/blob/master/Sources/Examples/1C8_3FilterExtractData/Script.py


keyboard
---------------
https://pypi.org/project/keyboard/


pathlib
---------------
https://docs.python.org/3/library/pathlib.html#pathlib.Path.is_file


win32clipboard
---------------
https://stackoverflow.com/questions/7050448/write-image-to-windows-clipboard-in-python-with-pil-and-win32clipboard

http://timgolden.me.uk/pywin32-docs/win32clipboard__IsClipboardFormatAvailable_meth.html


Popen
---------------
https://docs.python.org/3/library/subprocess.html#subprocess.Popen

Image module
---------------
https://pillow.readthedocs.io/en/stable/reference/Image.html

