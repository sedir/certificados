## Install

Install project dependencies

```
sudo apt install python3-tk
```

```
pip install -r requirements.txt
```

## Run

```
python3 main.py
```

## Use

![certificate generator photo](screenshot/screenshot.png)

**Certificate template**: accepts .JPG or .PNG formats.

**Input spreadsheet: accepts** .csv or .xlsx formats.

**Certificate text**: the spreadsheet information is filled in using `{}` or `{0}` corresponding to the column index, which represents data from the file line as you can see in the example image.

**Font**: Select a font that is installed on your system.

**Output folder**: folder where you will save the certificates.