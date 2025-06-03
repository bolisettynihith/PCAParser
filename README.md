# PCAParser

A Python script that parses and converts the .txt files from C:\Windows\appcompat\pca (Windows 11 22H2) to CSV files

# Dependencies

These are the required libraries needed to run this script.

+ argparse
+ csv
+ os
+ datetime

# Usage

This is a CLI based tool.

```
$ python PCAParser.py -i C:\Windows\appcompat\pca
```

![Usage](img\usage.png)

To view help:

```
$ python PCAParser.py -h
```

![help](img\help.png)

# References

+ https://aboutdfir.com/new-windows-11-pro-22h2-evidence-of-execution-artifact/
+ https://www.sygnia.co/blog/new-windows-11-pca-artifact/

# Code Inspiration

+ https://github.com/AndrewRathbun/PCAParser

Fyi, If you are in SOC/MDR, Andrew Rathbun's PowerShell script can be used to run directly on a remote Windows using the Remote execution capabilities provided by the EDR/XDR.

This script is prepared with my customizations, to help me during analysis of this artifact in investigations.