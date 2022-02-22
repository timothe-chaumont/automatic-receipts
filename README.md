# automatic_receipts

Python script that automates the generation and sending of receipts from orders data written in a Google Sheets.

## Steps to generate receipts automatically

- create a file named `.env` containing the same variables as described in `example.env`
- create a file named `associations_addresses.json` containing the same variables as described in `example.associations_addresses.json`.
- add a `.png` image of your association's logo in the current directory and call it `logo.png`. The image will be resized.
- install requirements with `pip install -r requirements.txt`
- run the main file `python main.py`. The options are :
  - `-a asso_name` to create the missing receipts of this association,
  - `-s` to print a summary of the orders that don't have any receipts yet.
