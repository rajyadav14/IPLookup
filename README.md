IP Lookup Tool
Overview
The IP Lookup Tool is a Python script that retrieves information about IP addresses from a spreadsheet using the IPWhois library. It processes multiple IP addresses concurrently, allowing for efficient data gathering and logging of errors.

Features
Reads a spreadsheet containing IP addresses.
Fetches details such as ASN (Autonomous System Number), ASN Name, State, and Country for each IP address.
Saves the results back into the spreadsheet.
Logs any errors encountered during the lookup process.
Requirements
Before running the script, make sure you have the following Python packages installed:

openpyxl: For reading and writing Excel files.
ipwhois: For fetching IP address details.
concurrent.futures: For handling concurrent processing.
You can install the required packages using pip:

bash
Copy code
pip install openpyxl ipwhois
How to Use
Prepare Your Spreadsheet: Create an Excel file with a list of IP addresses in the first column. Make sure to start from the second row, as the first row will be skipped.

Run the Script:

Open your terminal and navigate to the directory containing the script.
Execute the script using the following command:
bash
Copy code
python IPlookup.py
When prompted, enter the path to your spreadsheet.
Save the Results: After processing, you will be prompted to enter the path where you want to save the updated spreadsheet.

Error Logging
If any errors occur during the lookup, they will be logged in the error.log file located in the same directory as the script. You can check this log for details about any issues encountered with specific IP addresses.
