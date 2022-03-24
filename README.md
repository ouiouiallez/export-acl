# EXPORT ACL TO EXCEL FILE  

* * *

## About

This is a PowerShell script I made in order to audit the access rights in my company's shared folders.\
Thanks to the ImportExcel module by dfinke, the ACL are exported to one or more files (you choose), with a worksheet for every folder given in parameter.\
Written by Benoît Flache, 2022.  

![screenshot](https://raw.githubusercontent.com/ouiouiallez/ouiouiallez.github.io/master/content/pics/screenshot.JPG)

## Features

- Fetches the AD Security Groups, prints their names as well as the list of their members
- Customizable search depth
- Takes a single folder path or a file text with all paths to scan
- Adds a Table Style, you can change the default one with -style parameter
- Puts all scans in a single file with several worksheets, or into separate files with single worksheets.

## Prerequisites

- You need to have installed the RSAT tools on your computer, as this scripts uses some cmdlets included in those.
- You need of course access rights to the directories you want to scan.
- The ImportExcel module is automatically imported if you do not have it installed.

## How to use

### Parameters

`-out` is where you want to save the Excel file. For example : C:\document.xlsx

`-scan` is either :

- the path to the directory you want to scan
- the path to a txt file with the list of all the directories you want to scan
  
`-depth` is the recursive depth. Default : 1.

`-help` to print help and command examples

`-q` to disable output to console

`-style` to select the table style. Possibilities  are listed in file `help.txt`

`-split` if you want to scan several folders and have the results saved in different files.\
If you enabled this option, you have to give a folder and not a filename in the `-out` parameter.\
The files are named like SCANNED_DIRECTORY_NAME.xlsx.\
If there is already a file named like this in the folder, the results will be added as a second sheet in the existing file.

### txt file

To scan several folders in one shot, you can create a txt file containing all those folders separated by a line break
For example :

*folders.txt*

```text
K:\first\folder\to\scan
C:\second\folder
M:\
```

## Examples

```powershell
.\export-acl.ps1 -scan M:\path\to\directory -out C:\document.xlsx
.\export-acl.ps1 -scan C:\path\to\list.txt -out C:\document.xlsx
.\export-acl.ps1 -scan C:\path\to\list.txt -out C:\document.xlsx -depth 2
.\export-acl.ps1 -help
.\export-acl.ps1 -scan M:\path\to\directory -out C:\document.xlsx -style Medium3
.\export-acl.ps1 -scan C:\path\to\list.txt -out C:\directory -split -depth 0
```

## Links

[GitHub repository](https://github.com/ouiouiallez/export-acl)\
[ImportExcel GitHub repo](https://github.com/dfinke/ImportExcel)

## Improvements

I will try to improve this script, however if you have any questions or ideas on how to improve the code with new features or redesigning the functions, structure or in general code quality you are more than welcome :)
