# File Sorting PowerShell Script
## PowerShell Alphabetical Folder Sorting Algorithm
### Developed by Damon Murdoch ([@SirScrubbington](https://twitter.com/SirScrubbington)).

## Overview

This PowerShell script is designed to organize files in a folder into subfolders based on the initial characters of their filenames. It offers a flexible and customizable way to sort and manage large numbers of files quickly and efficiently.

## Features

- Sorts files into subfolders based on the initial characters of their filenames.
- Customizable options for folder naming, including prefix and suffix.
- Ability to combine or split folders based on file count.
- Recursive sorting for multiple layers of subfolders.
- WhatIf mode for previewing changes before applying them.
- Option to convert folder names to uppercase.

## Usage

1. Download the script (`Sort-Folder.ps1`) to your computer.

2. Open PowerShell and navigate to the directory containing the script.

3. Run the script with the desired parameters. Here's an example:

   ```powershell
   .\Sort-Folder.ps1 -Path "C:\YourFolderPath" -Force -Empty -Combine -Split -WhatIf -Recurse -Depth 2 -Upper -IncludeCount -Prefix "Prefix" -Suffix "Suffix" -Threshold 10

## Future Changes
A list of future planned fixes / improvements are listed below.

### Change Table
| Change Description | Priority |
| ------------------ | -------- |
| No planned changes | -        |

### Problems / Improvements
If you have any suggested improvements for this project or you encounter any issues, please feel free 
to open an issue or send me a message on twitter detailing the issue and how it can be replicated.

## Sponsor this Project
If you'd like to support this project and other future projects, 
please feel free to use the paypal domation link below.

https://www.paypal.com/paypalme/sirsc