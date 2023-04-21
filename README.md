# nova-data-processing-tool

Data processing toolkit for Nova Home Support (and maybe other residential care programs).

### Current capabilities:
* Drop irrelevant columns.

* Convert "Service 1 Auth Form ID" and "Service Provider" to desired formats.

* Correct "Check-In Date", "Check-In Time", "Check-Out Date", and "Check-Out Time".

* Check if "Staff Worked Duration" is consistent with "Staff Worked Duration (Minutes)" and check-in/check-out records. 
Save results to a new column named "Sanity Check".

### Directory:

  * README.md: This file.
  
  * process_data.ipynb: For testing code (no UI).
  
  * demo.py: **The actual data processor**.
  
  * testdata.xlsx: First report export data sent by Alison on April 20, 2023.
  
  * testdata(with issue).xlsx: First report export data with one additional row that has incorrect Check-Out Time.
  
  * data_processed.xlsx: Program output.

### Get Started

&nbsp;1. Download and install Anaconda from https://www.anaconda.com/download/

&nbsp;2. Download demo.py in this GitHub repository. 

&nbsp;3. In command line (Windows) or terminal (MacOS), type:

```shell
python3 .../file-path-on-your-computer/demo.py
```
