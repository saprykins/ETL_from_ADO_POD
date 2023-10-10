#!/bin/sh
source ./venv/bin/activate

ACCOUNT_KEY="1**"


rm -f *.csv
rm -f *.xlsx

python 3_main.py

az storage file upload --account-name azuredevopsblobstorage --account-key $ACCOUNT_KEY --share-name azuredevops-fs --source output2.xlsx
