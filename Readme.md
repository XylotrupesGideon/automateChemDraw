# Chemdraw automated file converter

I wrote this script to automatically convert a lot of ChemDraw cdx/cdxml files into image files by automating ChemDraw via the COM.
It requires an active installation of ChemDraw and Python.

It just takes a folder with ChemDraw files and a list of output formats.

## Usage

`python convert_cdxml.py "path/to/input/folder" [list of output format]`

e.g.
`python convert_cdxml.py "./Chemdraw" png svg`
