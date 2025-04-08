# from flask import Flask, render_template, request, redirect, send_file, flash
# import os
# from werkzeug.utils import secure_filename
# from Converter import json_to_excel, xml_to_excel, fetch_and_convert, database

# app = Flask(__name__)
# app.secret_key = 'supersecretkey'
# UPLOAD_FOLDER = 'uploads'
# CONVERTED_FOLDER = 'converted'

# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# database.initialize_db()

# @app.route('/')
# def index():
#     return render_template('index.html')

# @app.route('/upload/xml', methods=['GET', 'POST'])
# def upload_xml():
#     if request.method == 'POST':
#         file = request.files['file']
#         if file and file.filename.endswith('.xml'):
#             filename = secure_filename(file.filename)
#             path = os.path.join(UPLOAD_FOLDER, filename)
#             file.save(path)
#             out_path = xml_to_excel.convert_xml_to_excel(path, CONVERTED_FOLDER)
#             database.save_file_record(filename, 'xml', 'upload', out_path)
#             return send_file(out_path, as_attachment=True)
#         flash("Please upload a valid XML file.")
#     return render_template('xml_upload.html')

# @app.route('/upload/json', methods=['GET', 'POST'])
# def upload_json():
#     if request.method == 'POST':
#         file = request.files['file']
#         if file and file.filename.endswith('.json'):
#             filename = secure_filename(file.filename)
#             path = os.path.join(UPLOAD_FOLDER, filename)
#             file.save(path)
#             out_path = json_to_excel.convert_json_to_excel(path, CONVERTED_FOLDER)
#             database.save_file_record(filename, 'json', 'upload', out_path)
#             return send_file(out_path, as_attachment=True)
#         flash("Please upload a valid JSON file.")
#     return render_template('json_upload.html')

# @app.route('/fetch', methods=['GET', 'POST'])
# def fetch_url():
#     if request.method == 'POST':
#         url = request.form['url']
#         out_path, file_type, filename = fetch_and_convert.fetch_file_from_url(url, CONVERTED_FOLDER)
#         database.save_file_record(filename, file_type, 'url', out_path, url)
#         return send_file(out_path, as_attachment=True)
#     return render_template('fetch_url.html')

# if __name__ == '__main__':
#     app.run(debug=True)

# import requests

# # Tally local server URL (make sure Tally is running and ODBC/XML is enabled)
# url = "http://localhost:9000"

# # # XML request to fetch all Ledgers
# # xml_request = """
# # <ENVELOPE>
# #   <HEADER>
# #     <VERSION>1</VERSION>
# #     <TALLYREQUEST>Export</TALLYREQUEST>
# #     <TYPE>Collection</TYPE>
# #     <ID>Ledger Collection</ID>
# #   </HEADER>
# #   <BODY>
# #     <DESC>
# #       <STATICVARIABLES>
# #         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
# #       </STATICVARIABLES>
# #       <TDL>
# #         <TDLMESSAGE>
# #           <COLLECTION NAME="Ledger Collection" ISMODIFY="No">
# #             <TYPE>Ledger</TYPE>
# #           </COLLECTION>
# #         </TDLMESSAGE>
# #       </TDL>
# #     </DESC>
# #   </BODY>
# # </ENVELOPE>
# # """
# xml_request = """ 
# <ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>StockItems</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="StockItems" ISMODIFY="No">
#             <TYPE>StockItem</TYPE>
#             <FETCH>Name,Rate,CurrencyName,OpeningBalance,ClosingBalance,BaseUnits,HSNCODE,HSNDESCRIPTION</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE> """

# # Recommended headers
# headers = {
#     "Content-Type": "application/xml",  # or "text/xml"
# }

# # POST the XML to Tally
# response = requests.post(url, data=xml_request.encode('utf-8'), headers=headers)

# # Print the XML response from Tally
# print(response.text)


# import requests
# import xml.etree.ElementTree as ET
# import pandas as pd

# url = "http://localhost:9000"

# xml_request = """<ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>StockItems</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="StockItems" ISMODIFY="No">
#             <TYPE>StockItem</TYPE>
#             <FETCH>Name,Rate,CurrencyName,OpeningBalance,ClosingBalance,BaseUnits,HSNCODE,HSNDESCRIPTION</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>"""

# headers = {"Content-Type": "application/xml"}

# response = requests.post(url, data=xml_request.encode('utf-8'), headers=headers)

# # Parse XML response
# root = ET.fromstring(response.text)
# data = []

# for stock_item in root.findall(".//STOCKITEM"):
#     data.append({
#         "Product Name": stock_item.findtext("NAME", default=""),
#         "Sales Price": stock_item.findtext("RATE", default=""),
#         "Cost Currency": stock_item.findtext("CURRENCYNAME", default=""),
#         "Quantity On Hand": stock_item.findtext("CLOSINGBALANCE", default=""),
#         "Forecasted Quantity": "0.00",  # Default value, not in Tally
#         "Cost": stock_item.findtext("OPENINGBALANCE", default=""),
#         "Unit of Measure": stock_item.findtext("BASEUNITS", default=""),
#         "HSN/SAC Code": stock_item.findtext("HSNCODE", default=""),
#         "HSN/SAC Description": stock_item.findtext("HSNDESCRIPTION", default="")
#         # Add more fields here as needed
#     })

# # Convert to DataFrame
# df = pd.DataFrame(data)

# # Optional: Save to Excel
# df.to_excel("tally_products.xlsx", index=False)

# print(df.head())


# import re
# import xml.etree.ElementTree as ET
# import requests
# import pandas as pd

# # Send XML request to Tally
# url = "http://localhost:9000"
# xml_request = """<ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>StockItems</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="StockItems" ISMODIFY="No">
#             <TYPE>StockItem</TYPE>
#             <FETCH>Name,Rate,CurrencyName,OpeningBalance,ClosingBalance,BaseUnits,HSNCODE,HSNDESCRIPTION</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>"""

# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# # CLEANING STEP 👇 to fix invalid characters
# def clean_xml(xml_str):
#     # Remove invalid XML characters using regex
#     xml_str = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '', xml_str)
#     # Replace unescaped & with &amp; unless already part of an entity
#     xml_str = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)', '&amp;', xml_str)
#     return xml_str

# cleaned_xml = clean_xml(response.text)

# # Now safely parse the XML
# try:
#     root = ET.fromstring(cleaned_xml)
# except ET.ParseError as e:
#     print("Error parsing XML:", e)
#     exit()

# # Extract data
# data = []
# for stock_item in root.findall(".//STOCKITEM"):
#     data.append({
#         "Product Name": stock_item.findtext("NAME", default=""),
#         "Sales Price": stock_item.findtext("RATE", default=""),
#         "Cost Currency": stock_item.findtext("CURRENCYNAME", default=""),
#         "Quantity On Hand": stock_item.findtext("CLOSINGBALANCE", default=""),
#         "Forecasted Quantity": "0.00",  # Placeholder
#         "Cost": stock_item.findtext("OPENINGBALANCE", default=""),
#         "Unit of Measure": stock_item.findtext("BASEUNITS", default=""),
#         "HSN/SAC Code": stock_item.findtext("HSNCODE", default=""),
#         "HSN/SAC Description": stock_item.findtext("HSNDESCRIPTION", default="")
#     })

# # Convert to DataFrame and export
# df = pd.DataFrame(data)
# df.to_excel("tally_products_cleaned.xlsx", index=False)

# print(df.head())


# import re
# import xml.etree.ElementTree as ET
# import requests
# import pandas as pd

# # --------------------- Step 1: Send XML request to Tally ---------------------
# url = "http://localhost:9000"

# xml_request = """<ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>StockItems</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="StockItems" ISMODIFY="No">
#             <TYPE>StockItem</TYPE>
#             <FETCH>Name,Rate,CurrencyName,OpeningBalance,ClosingBalance,BaseUnits,HSNCODE,HSNDESCRIPTION</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>"""

# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# # --------------------- Step 2: Save raw XML to file ---------------------
# with open("raw_stock_items.xml", "w", encoding="utf-8") as file:
#     file.write(response.text)

# # --------------------- Step 3: Clean the XML ---------------------
# def clean_xml(xml_str):
#     # Remove invalid XML characters
#     xml_str = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '', xml_str)
#     # Escape unescaped ampersands
#     xml_str = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)', '&amp;', xml_str)
#     return xml_str

# cleaned_xml = clean_xml(response.text)

# # Save cleaned XML
# with open("cleaned_stock_items.xml", "w", encoding="utf-8") as file:
#     file.write(cleaned_xml)

# # --------------------- Step 4: Parse cleaned XML ---------------------
# try:
#     root = ET.fromstring(cleaned_xml)
# except ET.ParseError as e:
#     print("❌ Error parsing XML:", e)
#     exit()

# # --------------------- Step 5: Extract Data ---------------------
# data = []
# for stock_item in root.findall(".//STOCKITEM"):
#     data.append({
#         "Product Name": stock_item.findtext("NAME", default=""),
#         "Sales Price": stock_item.findtext("RATE", default=""),
#         "Cost Currency": stock_item.findtext("CURRENCYNAME", default=""),
#         "Quantity On Hand": stock_item.findtext("CLOSINGBALANCE", default=""),
#         "Forecasted Quantity": "0.00",  # Static placeholder
#         "Cost": stock_item.findtext("OPENINGBALANCE", default=""),
#         "Unit of Measure": stock_item.findtext("BASEUNITS", default=""),
#         "HSN/SAC Code": stock_item.findtext("HSNCODE", default=""),
#         "HSN/SAC Description": stock_item.findtext("HSNDESCRIPTION", default="")
#     })

# # --------------------- Step 6: Save to Excel ---------------------
# df = pd.DataFrame(data)
# df.to_excel("tally_products_cleaned.xlsx", index=False)

# print("✅ Data exported to 'tally_products_cleaned.xlsx'")
# print(df.head())


import re
import xml.etree.ElementTree as ET
import requests
import pandas as pd

# --------------------- Step 1: Send XML request to Tally ---------------------
url = "http://localhost:9000"

xml_request = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>StockItems</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="StockItems" ISMODIFY="No">
            <TYPE>StockItem</TYPE>
            <FETCH>Name,Rate,CurrencyName,OpeningBalance,ClosingBalance,BaseUnits,HSNCODE,HSNDESCRIPTION</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# --------------------- Step 2: Save raw XML to file ---------------------
with open("raw_stock_items1.xml", "w", encoding="utf-8") as file:
    file.write(response.text)

# --------------------- Step 3: Clean the XML ---------------------
def clean_xml(xml_str):
    # Remove characters not allowed by XML 1.0 standard
    def is_valid_xml_char(char):
        codepoint = ord(char)
        return (
            codepoint == 0x9 or
            codepoint == 0xA or
            codepoint == 0xD or
            (0x20 <= codepoint <= 0xD7FF) or
            (0xE000 <= codepoint <= 0xFFFD) or
            (0x10000 <= codepoint <= 0x10FFFF)
        )

    cleaned = ''.join(c for c in xml_str if is_valid_xml_char(c))

    # Escape unescaped ampersands
    cleaned = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)', '&amp;', cleaned)

    return cleaned

cleaned_xml = clean_xml(response.text)

# Save cleaned XML
with open("cleaned_stock_items.xml", "w", encoding="utf-8") as file:
    file.write(cleaned_xml)

# --------------------- Step 4: Parse cleaned XML ---------------------
try:
    root = ET.fromstring(cleaned_xml)
except ET.ParseError as e:
    print("❌ Error parsing XML:", e)
    exit()

# --------------------- Step 5: Extract Data ---------------------
data = []
for stock_item in root.findall(".//STOCKITEM"):
    data.append({
        "Product Name": stock_item.findtext("NAME", default=""),
        "Sales Price": stock_item.findtext("RATE", default=""),
        "Cost Currency": stock_item.findtext("CURRENCYNAME", default=""),
        "Quantity On Hand": stock_item.findtext("CLOSINGBALANCE", default=""),
        "Forecasted Quantity": "0.00",  # Static placeholder
        "Cost": stock_item.findtext("OPENINGBALANCE", default=""),
        "Unit of Measure": stock_item.findtext("BASEUNITS", default=""),
        "HSN/SAC Code": stock_item.findtext("HSNCODE", default=""),
        "HSN/SAC Description": stock_item.findtext("HSNDESCRIPTION", default="")
    })

# --------------------- Step 6: Save to Excel ---------------------
df = pd.DataFrame(data)
df.to_excel("tally_products_cleaned.xlsx", index=False)

print("✅ Data exported to 'tally_products_cleaned.xlsx'")
print(df.head())
