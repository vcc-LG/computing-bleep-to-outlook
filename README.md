# Computing Bleep Rota Parser

About
=====
A simple Python tool to read in the computing bleep rota Excel file and create Outlook events for your rota slots.

Usage
=====
```
pip install -r requirements.txt
```

Modify `config.py` to your specifications:

```
file_path = <full path to rota Excel file>
sheet_label = 'Jan17-Mar17'
user = <your initials> 
```

Run using 

```
python main.py
```

