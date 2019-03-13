# PythonExcelReader
[![Build Status](https://travis-ci.org/adamrees89/PythonExcelReader.svg?branch=master)](https://travis-ci.org/adamrees89/PythonExcelReader) [![Known Vulnerabilities](https://snyk.io/test/github/adamrees89/PythonExcelReader/badge.svg?targetFile=requirements.txt)](https://snyk.io/test/github/adamrees89/PythonExcelReader?targetFile=requirements.txt)
[![Coverage Status](https://coveralls.io/repos/github/adamrees89/PythonExcelReader/badge.svg?branch=master)](https://coveralls.io/github/adamrees89/PythonExcelReader?branch=master)

Python scripts to read an excel file and print the contents to a SQLlite3 database.  Future work involves re-writing that excel file elsewhere

## templateReader.py

Reads an excel file, then creates an SQLlite3 database with the values and formatting

## CreateSSExcelDoc.py

This crates an excel document based on the database created by templateReader.py
