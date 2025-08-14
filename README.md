<h2>ExcelFile-Filter</h2>
<br>
<b>ExcelFile-Filter</b> - a script designed to simplify searching in Excel files. 
<br><br>
:heavy_minus_sign:<b>Crucial Features</b>:heavy_minus_sign:
<ul>
    <li>All interactions with Excel documents are provided through user-friendly GUI.</li>
    <li>Handles user-generated exceptions gracefully with informative message boxes.</li>
    <li>The filtered output is saved as a professionally formatted Excel document.</li>
    <li>The program is designed to be easy to use and offers simple maintenance.</li>
</ul>
<br>
:eight_spoked_asterisk:<b>Tech Stack</b>:eight_spoked_asterisk:<br>
Python, Tkinter, Openpyxl, Datetime lib.
<br>
<hr>
<b>Project deployment on Windows</b>
<br><br>

```bash
git clone https://github.com/735Andrew/ExcelFile-Filter
cd ExcelFile-Filter
python -m venv venv 
venv\Scripts\activate
(venv) pip install -r requirements.txt
    
(venv) python gui.py                      # Opens the GUI             
     
```
<hr>
<b>Project deployment on Linux</b>
<h6>Prerequisites:</h6>

```bash
sudo apt-get -y install python3-tk
```
<br>

```bash
git clone https://github.com/735Andrew/ExcelFile-Filter
cd ExcelFile-Filter
python3 -m venv venv 
source venv/bin/activate 
(venv) pip install -r requirements.txt
   		
(venv) python3 gui.py                     # Opens the GUI               
    
```