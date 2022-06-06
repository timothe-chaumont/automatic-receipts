# automatic_receipts

Python script that automates the generation and sending of receipts from orders data written in a Google Sheets.

## Steps to generate receipts automatically

1. create a file named _.env_ containing the same variables as described in [example.env](/example.env).
2. create a file named _associations_addresses.json_ containing the same variables as described in [example.associations_addresses.json](example.associations_addresses.json).
3. add a _.png_ image of your association's logo in the current directory and call it _logo.png_. The image will be resized.
4. install requirements with `pip install -r requirements.txt`
5. run `process_all_orders.py` file. It will show you a summary of what will be done, then ask you to confirm.
