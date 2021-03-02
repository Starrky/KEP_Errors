# KEP_Errors
 
When you run batch file script will download csv files of tables taken from SQL database using SQLCMD then process them and present summary in Results/Error_status.xlsx.


# Installation

**In order to use this tool, you need to install python3 and chrome browser and keep it updated**: https://www.python.org/downloads/ and choose newest version of the installers for your OS**(select "Add to path" checkbox during installation)**.

![alt text](https://i.imgur.com/06EspWQ.png)

Use virtual environment, in project folder:

Create and activate venv:
```
python -m venv c:\path\to\myenv
c:\path\to\myenv\Scripts\activate.bat
```

Then install requirements
```
pip install -r requirements.txt
```

You will also need to have SQLCDM installed: https://docs.microsoft.com/en-us/sql/tools/sqlcmd-utility?view=sql-server-ver15

**For ubuntu**  you have to:

```
pip3 install -r requirements.txt
```

# Usage
Make sure you are connected to comapny VPN, then simply run batch file "Run.bat" and wait for it to end, then check the results in Results\Errors_status.xlsx
