# automatic_receipts

Python script that automates the generation and sending of receipts from orders data written in a Google Sheets.

_It was initially designed for the student association [CSDesign](https://csdesign.cs-campus.fr/)_.

## To generate receipts automatically

1. **Add** the following **files** to the root folder*:
    - _.env_ containing the same variables as described in [example.env](/example.env).
    -  _associations_addresses.json_ containing the same variables as described in [example.associations_addresses.json](example.associations_addresses.json).
    -  A _.png_ image of your association's logo in the current directory and call it _logo.png_. The image will be resized.
    -  _credentials.json_ which is provided by Google Cloud when requesting an access to Google Sheets API.
1. **Install** the **required packages** by running `pip install -r requirements.txt` in a terminal (at the same level as the _requirements.txt_ file)
1. **Run the script** `process_all_orders.py`. This will first show you a summary of what will be done and ask for your confirmation. It will then process all the receipts for which it is possible.

* _If you are a member of CSDesign, you can ask a previous tresurer for those files_

## How does it work ?

![Architecture diagram](/architecture_diagram.png)
