# datasheet_automation
This script automates the creation of the Master Equipment datasheet file and populates the sheet with the SysCAD inputs

# 📊 Master Equipment Sheet Generator

This Streamlit web app allows internal users to generate a master equipment data sheet by uploading a standardized Excel workbook (`Datasheets.xlsm`). The app automatically processes and formats the data across equipment sheets and provides a timestamped Excel file ready for download.

---

## 🚀 Features

- ✅ Upload Excel file (`.xlsm`) with multiple equipment sheets
- ✅ Extracts parameters, units, and categories from each sheet
- ✅ Groups parameters under:
  - SysCAD Inputs
  - Engineering Inputs
  - Lab/Pilot Inputs
  - Project Constant
  - Vendor Inputs
- ✅ Merges category cells
- ✅ Auto-sizes columns based on content
- ✅ Applies clean border formatting to all data cells
- ✅ Generates output with date+time stamped filename
- ✅ Downloadable masterdata sheet output file from the UI
- ✅ Uplaod Excel file with syscad input parameters in a stream table
- ✅ Display the list of SysCAD inputs from the master
- ✅ Let the user to map it with the respective parameters in SysCAD
- ✅ Shows warning for equipment missing in the streamtable
- ✅ Saves the mapping and populates the values accordingly
- ✅ Rounds up the values to 2 decimal places and update units if needed
- ✅ Downlaodable populated output file from the UI

---


