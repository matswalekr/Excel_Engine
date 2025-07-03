# Excel Engine

## Description

A Python File used to interface with Excel. 

It allows to manipulate individual cells, interpret the file as a pandas dataframe or evaluate cells using Python.

## Usage

To use the Engine, it is recommended to use the open_excel function, which is a factory for the different modes to open a file.

An excel file may be opened in any of the following modes:
- "r": **read-mode**, No changes may be made to individual cells
- "m": **modify-mode**, Changes may be made to the excel sheet, but these can't be saved. Use this mode to run simulations or model using different inputs.
- "w": **write-mode**, Changes may be made to the Excel and these can be saved in different locations.

Use the Engine as follows:
with open_excel(path = r'some_path.xlsx", mode = "w") as xls: 

## Dependencies
This file extends different libraries.

It depends on the following ones:
- openpyxl
- pycel
- pandas

## Work to be done
The Excel libraries that this Engine is built on top of have issues when working with different Excel versions. As a result, the files may show warnings that they are corrupted. However, they can usually still be used.

In addition, graphs used in templates which are then modified in using the Engine may change or behave unexpectedly.

The usage of VBA macros has not been extensively tested. However, there exist ways to execute them.