# automatic_receipts

Python script that automates the generation and sending of receipts from orders data written in a Google Sheets.

## Steps to generate receipts automatically

1. create a file named *.env* containing the same variables as described in [example.env](/example.env).
2. create a file named *associations_addresses.json* containing the same variables as described in [example.associations_addresses.json](example.associations_addresses.json).
3. add a *.png* image of your association's logo in the current directory and call it *logo.png*. The image will be resized.
4. install requirements with `pip install -r requirements.txt`
5. run the main file `python main.py`. The options are :
   - `-a asso_name` to create the missing receipts of this association,
   - `-s` to print a summary of the orders that don't have any receipts yet.
