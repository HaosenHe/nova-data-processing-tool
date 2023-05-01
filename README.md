# nova-data-processing-tool

Data processing toolkit for Nova Home Support (and other residential care programs).

Version 0.4

### Current Capabilities:
* Drop irrelevant columns.

* Convert "Service 1 Description (Code)" and "Service Provider" to desired formats.

* Correct "Check-In Date", "Check-In Time", "Check-Out Date", and "Check-Out Time".

* Check if "Staff Worked Duration" is consistent with "Staff Worked Duration (Minutes)" and check-in/check-out records. 
Save results to a new column named "Sanity Check".

* Process payroll for managers and non-managers; Output Total Gross Wages,	Total Hours Worked, and	Accured Time Off for the pay period. 

### File Directory:

  * README.md: This file.
    
  * demo.py: **The actual data processor**.
  
  * helpers.py: Helper functions (must be in the same folder as demo.py).
  
  * requirements.txt: Dependencies.


### Get Started

&nbsp;1. Download and install Anaconda from https://www.anaconda.com/download/

&nbsp;2. Download noval-data-processing-tool from this GitHub repository. 

&nbsp;3. In command line (Windows) or terminal (MacOS), go to the noval-data-processing-tool directory, type:

```shell
pip install -r requirements.txt
```

&nbsp; to install dependencies. Once dependencies are installed, type the following to open the program:

```shell
python3 demo.py
```

&nbsp;4. You should now see the following user interface:

<img width="1274" alt="image" src="https://user-images.githubusercontent.com/29806214/235276733-6665ad76-508a-4b72-9212-4646e7747d48.png">

&nbsp;5. When everything is done correctly, you should see:

<img width="1274" alt="image" src="https://user-images.githubusercontent.com/29806214/235276722-f509baa5-b5a2-49b4-b492-9af84ff9169d.png">
