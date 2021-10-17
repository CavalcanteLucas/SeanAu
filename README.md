## Instructions of use

### Prerequisites:

- Create and activate virtual environment, then install the project dependencies:
```
virtualenv venv
source venv/bin/activate
pip install -r requirements.txt
```

- Make sure to have the list `clickid.xlsx` data in the same directory of the script `main.py`.

### Running the script

- Run the script with the command:

```
python main.py
```

By default, the **verbose** mode is `True`. So it will log out every instance of `clickid` processed. If this should be supressed, run the script with the command:

```
python main.py --verbse false
```

- The script output will be recorded into the file `output.xlsx`.