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

# import re
# import xml.etree.ElementTree as ET
# import requests
# import pandas as pd

# # --------------------- Step 1: Send XML request to Tally ---------------------
# url = "http://localhost:9000"

# xml_request = """
# <ENVELOPE>
#   <HEADER>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ & Co. (2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="AllStockItems">
#             <TYPE>StockItem</TYPE>
#             <FETCH>
#               NAME,
#               BASEUNITS,
#               COSTINGMETHOD,
#               STANDARDPRICE,
#               OPENINGBALANCE,
#               CLOSINGBALANCE,
#               RATE,
#               TAXCLASSIFICATIONNAME,
#               HSN/SAC,
#               HSNCODE,
#               HSNDESCRIPTION,
#               GSTAPPLICABLE,
#               GSTTYPE,
#               CATEGORY,
#               PARTNUMBER,
#               ISBATCHWISEON,
#               ISINVENTORYON,
#               ISPRICEINCLUSIVEOFTAX,
#               ISUPDATINGREALTIME,
#               USERDEFINEDFIELDLIST.LIST,
#               GUID
#             </FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>

# """


# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# # --------------------- Step 2: Save raw XML to file ---------------------
# with open("raw_stock_items3.xml", "w", encoding="utf-8") as file:
#     file.write(response.text)

# # --------------------- Step 3: Clean the XML ---------------------
# def clean_xml(xml_str):
#     # Remove characters not allowed by XML 1.0 standard
#     def is_valid_xml_char(char):
#         codepoint = ord(char)
#         return (
#             codepoint == 0x9 or
#             codepoint == 0xA or
#             codepoint == 0xD or
#             (0x20 <= codepoint <= 0xD7FF) or
#             (0xE000 <= codepoint <= 0xFFFD) or
#             (0x10000 <= codepoint <= 0x10FFFF)
#         )

#     cleaned = ''.join(c for c in xml_str if is_valid_xml_char(c))

#     # Escape unescaped ampersands
#     cleaned = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)', '&amp;', cleaned)

#     return cleaned

# cleaned_xml = clean_xml(response.text)

# # Save cleaned XML
# with open("cleaned_stock_items6.xml", "w", encoding="utf-8") as file:
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
# df.to_excel("tally_products_cleaned5.xlsx", index=False)

# print("✅ Data exported to 'tally_products_cleaned.xlsx'")
# print(df.head())



# import re
# import requests

# # --------------------- Step 1: Send XML request to Tally ---------------------
# url = "http://localhost:9000"

# # Corrected XML request with proper company name and escaped ampersand
# xml_request = """
# <ENVELOPE>
#   <HEADER>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="AllStockItems">
#             <TYPE>StockItem</TYPE>
#             <FETCH>
#               NAME,
#               BASEUNITS,
#               COSTINGMETHOD,
#               STANDARDPRICE,
#               OPENINGBALANCE,
#               CLOSINGBALANCE,
#               RATE,
#               TAXCLASSIFICATIONNAME,
#               HSN/SAC,
#               HSNCODE,
#               HSNDESCRIPTION,
#               GSTAPPLICABLE,
#               GSTTYPE,
#               CATEGORY,
#               PARTNUMBER,
#               ISBATCHWISEON,
#               ISINVENTORYON,
#               ISPRICEINCLUSIVEOFTAX,
#               ISUPDATINGREALTIME,
#               USERDEFINEDFIELDLIST.LIST,
#               GUID
#             </FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# # Send request to Tally
# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# # --------------------- Step 2: Save raw XML response to file ---------------------
# with open("raw_stock_items4.xml", "w", encoding="utf-8") as file:
#     file.write(response.text)

# print("✅ Raw XML response saved to 'raw_stock_items4.xml'")


# import requests

# url = "http://localhost:9000"

# # XML request to fetch all Ledgers
# xml_request = """
# <ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>Ledger Collection</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="Ledger Collection" ISMODIFY="No">
#             <TYPE>Ledger</TYPE>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# response = requests.post(url, data=xml_request)
# print(response.text)






# import re
# import requests

# # --------------------- Step 1: Send XML request to Tally ---------------------
# url = "http://localhost:9000"

# # Corrected XML request with properly escaped ampersand in company name
# xml_request = """
# <ENVELOPE>
#   <HEADER>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="AllStockItems">
#             <TYPE>StockItem</TYPE>
#             <FETCH>
#               NAME,
#               BASEUNITS,
#               COSTINGMETHOD,
#               STANDARDPRICE,
#               OPENINGBALANCE,
#               CLOSINGBALANCE,
#               RATE,
#               TAXCLASSIFICATIONNAME,
#               HSN/SAC,
#               HSNCODE,
#               HSNDESCRIPTION,
#               GSTAPPLICABLE,
#               GSTTYPE,
#               CATEGORY,
#               PARTNUMBER,
#               ISBATCHWISEON,
#               ISINVENTORYON,
#               ISPRICEINCLUSIVEOFTAX,
#               ISUPDATINGREALTIME,
#               USERDEFINEDFIELDLIST.LIST,
#               GUID
#             </FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# # --------------------- Step 2: Send request and save XML ---------------------
# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# # Save raw XML to file
# with open("raw_stock_items5.xml", "w", encoding="utf-8") as file:
#     file.write(response.text)

# print("✅ Raw XML response saved to 'raw_stock_items4.xml'")


# import re
# import requests

# # --------------------- Step 1: Send XML request to Tally ---------------------
# url = "http://localhost:9000"

# # XML request with properly escaped ampersand (&amp;) in company name
# xml_request = """
# <ENVELOPE>
#   <HEADER>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="AllStockItems">
#             <TYPE>StockItem</TYPE>
#             <FETCH>NAME</FETCH>
#             <FETCH>BASEUNITS</FETCH>
#             <FETCH>COSTINGMETHOD</FETCH>
#             <FETCH>STANDARDPRICE</FETCH>
#             <FETCH>OPENINGBALANCE</FETCH>
#             <FETCH>CLOSINGBALANCE</FETCH>
#             <FETCH>RATE</FETCH>
#             <FETCH>TAXCLASSIFICATIONNAME</FETCH>
#             <FETCH>HSN/SAC</FETCH>
#             <FETCH>HSNCODE</FETCH>
#             <FETCH>HSNDESCRIPTION</FETCH>
#             <FETCH>GSTAPPLICABLE</FETCH>
#             <FETCH>GSTTYPE</FETCH>
#             <FETCH>CATEGORY</FETCH>
#             <FETCH>PARTNUMBER</FETCH>
#             <FETCH>ISBATCHWISEON</FETCH>
#             <FETCH>ISINVENTORYON</FETCH>
#             <FETCH>ISPRICEINCLUSIVEOFTAX</FETCH>
#             <FETCH>ISUPDATINGREALTIME</FETCH>
#             <FETCH>USERDEFINEDFIELDLIST.LIST</FETCH>
#             <FETCH>GUID</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# # --------------------- Step 2: Send request and save XML ---------------------
# try:
#     response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})
#     response.raise_for_status()  # Raises an error if response code is not 200 OK
#     with open("raw_stock_items7.xml", "w", encoding="utf-8") as file:
#         file.write(response.text)
#     print("✅ Raw XML response saved to 'raw_stock_items7.xml'")
# except requests.exceptions.RequestException as e:
#     print("❌ Error while sending request to Tally:", e)


# import requests

# url = "http://localhost:9000"  # Tally ODBC default port

# xml_request = """
# <ENVELOPE>
#   <HEADER>
#     <TALLYREQUEST>Export Data</TALLYREQUEST>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="AllLedgers">
#             <TYPE>Ledger</TYPE>
#             <FETCH>NAME</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# print(response.text)

# import requests

# url = "http://localhost:9000"  # Change if Tally uses a different port

# xml = """
# <ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>Ledger</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="Ledgers" ISMODIFY="No">
#             <TYPE>Ledger</TYPE>
#             <FETCH>NAME,PARENT,OPENINGBALANCE</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# response = requests.post(url, data=xml.encode('utf-8'), headers={"Content-Type": "application/xml"})

# print(response.text)






# import requests
# import xml.etree.ElementTree as ET
# import pandas as pd
# import os

# # Ask user for filename
# filename = input("Enter the base filename (without extension): ").strip()

# # Define paths and make sure folders exist
# upload_dir = r"D:\xml-json-converter\uploads"
# converted_dir = r"D:\xml-json-converter\converted"
# os.makedirs(upload_dir, exist_ok=True)
# os.makedirs(converted_dir, exist_ok=True)

# xml_path = os.path.join(upload_dir, f"{filename}.xml")
# excel_path = os.path.join(converted_dir, f"{filename}.xlsx")

# # XML request
# url = "http://localhost:9000"
# xml = """
# <ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>Ledger</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="Ledgers" ISMODIFY="No">
#             <TYPE>Ledger</TYPE>
#             <FETCH>NAME,PARENT,OPENINGBALANCE</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# # Step 1: Send request
# response = requests.post(url, data=xml.encode('utf-8'), headers={"Content-Type": "application/xml"})

# # Step 2: Save XML response
# with open(xml_path, "w", encoding="utf-8") as f:
#     f.write(response.text)
# print(f"✅ XML saved as: {xml_path}")

# # Step 3: Parse XML
# tree = ET.parse(xml_path)
# root = tree.getroot()

# ledgers = []
# for ledger in root.iter("LEDGER"):
#     data = {
#         "NAME": ledger.findtext("NAME", default=""),
#         "PARENT": ledger.findtext("PARENT", default=""),
#         "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
#     }
#     ledgers.append(data)

# # Step 4: Convert to Excel
# if ledgers:
#     df = pd.DataFrame(ledgers)
#     df.to_excel(excel_path, index=False)
#     print(f"✅ Excel saved as: {excel_path}")
# else:
#     print("⚠️ No ledger data found in XML.")



# import os
# import re
# import html
# import pandas as pd
# import xml.etree.ElementTree as ET

# # ---------- Step 1: User Input ----------
# base_filename = input("Enter base filename (without extension): ").strip()
# raw_file = f"D:/xml-json-converter/uploads/{base_filename}.xml"
# clean_dir = r"D:\xml-json-converter\Clean_xml"
# excel_dir = r"D:\xml-json-converter\converted"
# os.makedirs(clean_dir, exist_ok=True)
# os.makedirs(excel_dir, exist_ok=True)

# utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
# cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
# fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
# excel_file = os.path.join(excel_dir, f"{base_filename}.xlsx")

# # ---------- Step 2: Convert to UTF-8 ----------
# with open(raw_file, "r", encoding="utf-16", errors="ignore") as f:
#     content = f.read()

# with open(utf8_file, "w", encoding="utf-8") as f:
#     f.write(content)

# print(f"✅ UTF-8 conversion done: {utf8_file}")

# # ---------- Step 3: Clean XML (Control Characters + Entities + & Fix) ----------
# def clean_xml(file_path, output_file):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Remove non-printable characters
#     content = re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', content)

#     # Decode escaped HTML characters
#     content = html.unescape(content)

#     # Fix standalone '&'
#     content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Cleaned XML saved: {output_file}")
#     return output_file

# clean_xml(utf8_file, cleaned_file)

# # ---------- Step 4: Fix Unescaped '<' or '>' in Text ----------
# def fix_invalid_xml(file_path, output_file):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Fix unescaped < and > inside tags like <NAME>
#     content = re.sub(r'(<NAME>[^<]*?)(<)([^>]*?)(</NAME>)', r'\1&lt;\3\4', content)

#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Final XML fixed and saved: {output_file}")
#     return output_file

# fix_invalid_xml(cleaned_file, fixed_file)

# # ---------- Step 5: Convert Final XML to Excel ----------
# def xml_to_excel(xml_path, excel_path):
#     tree = ET.parse(xml_path)
#     root = tree.getroot()

#     data = []
#     for ledger in root.iter("LEDGER"):
#         record = {
#             "NAME": ledger.findtext("NAME", default=""),
#             "PARENT": ledger.findtext("PARENT", default=""),
#             "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
#         }
#         data.append(record)

#     if data:
#         df = pd.DataFrame(data)
#         df.to_excel(excel_path, index=False)
#         print(f"✅ Excel created: {excel_path}")
#     else:
#         print("⚠️ No LEDGER data found in the XML.")

# xml_to_excel(fixed_file, excel_file)





# import os
# import re
# import html
# import requests

# # --------- Ask for base filename ----------
# base_name = input("Enter base filename (without extension): ").strip()

# # Define folder paths
# upload_dir = "D:/xml-json-converter/uploads"
# clean_dir = "D:/xml-json-converter/Clean_xml"
# converted_dir = "D:/xml-json-converter/converted"

# # Define file paths
# raw_file = os.path.join(upload_dir, base_name + ".xml")
# cleaned_file = os.path.join(clean_dir, base_name + "_cleaned.xml")

# # Create folders if not exist
# os.makedirs(upload_dir, exist_ok=True)
# os.makedirs(clean_dir, exist_ok=True)
# os.makedirs(converted_dir, exist_ok=True)

# # --------- Step 1: Fetch XML from Tally and Save ---------
# url = "http://localhost:9000"
# xml_request = f"""
# <ENVELOPE>
#   <HEADER>
#     <VERSION>1</VERSION>
#     <TALLYREQUEST>Export</TALLYREQUEST>
#     <TYPE>Collection</TYPE>
#     <ID>Ledger</ID>
#   </HEADER>
#   <BODY>
#     <DESC>
#       <STATICVARIABLES>
#         <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#         <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#       </STATICVARIABLES>
#       <TDL>
#         <TDLMESSAGE>
#           <COLLECTION NAME="Ledgers" ISMODIFY="No">
#             <TYPE>Ledger</TYPE>
#             <FETCH>NAME,PARENT,OPENINGBALANCE</FETCH>
#           </COLLECTION>
#         </TDLMESSAGE>
#       </TDL>
#     </DESC>
#   </BODY>
# </ENVELOPE>
# """

# response = requests.post(url, data=xml_request.encode('utf-8'), headers={"Content-Type": "application/xml"})

# with open(raw_file, "w", encoding="utf-8") as f:
#     f.write(response.text)

# print(f"✅ Raw XML saved to: {raw_file}")

# # --------- Step 2: Clean the XML and save to /Clean_xml ----------
# def clean_xml(file_path, cleaned_path):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Remove invalid XML characters
#     content = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', content)

#     # Decode escaped characters (&amp;, etc.)
#     content = html.unescape(content)

#     # Fix stray '&' characters
#     content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

#     with open(cleaned_path, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Cleaned XML saved to: {cleaned_path}")

# # Run cleaning
# clean_xml(raw_file, cleaned_file)



# import os
# import re
# import html
# import pandas as pd
# import xml.etree.ElementTree as ET

# # ---------- Step 1: User Input ----------
# base_filename = input("Enter base filename (without extension): ").strip()

# # Folder paths
# upload_dir = "D:\xml-json-converter\uploads"
# clean_dir = "D:\xml-json-converter\Clean_xml"
# converted_dir = "D:\xml-json-converter\converted"

# # Create folders if not exist
# os.makedirs(upload_dir, exist_ok=True)
# os.makedirs(clean_dir, exist_ok=True)
# os.makedirs(converted_dir, exist_ok=True)

# # File paths
# raw_file = os.path.join(upload_dir, base_filename + ".xml")
# utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
# cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
# fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
# excel_file = os.path.join(converted_dir, f"{base_filename}.xlsx")

# # ---------- Step 2: Convert to UTF-8 ----------
# with open(raw_file, "r", encoding="utf-16", errors="ignore") as f:
#     content = f.read()

# with open(utf8_file, "w", encoding="utf-8") as f:
#     f.write(content)

# print(f"✅ UTF-8 conversion done: {utf8_file}")

# # ---------- Step 3: Clean XML ----------
# def clean_xml(file_path, output_file):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Remove non-printable characters
#     content = re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', content)

#     # Decode HTML entities
#     content = html.unescape(content)

#     # Fix standalone '&' not part of entities
#     content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Cleaned XML saved: {output_file}")
#     return output_file

# clean_xml(utf8_file, cleaned_file)

# # ---------- Step 4: Fix Unescaped '<' or '>' ----------
# def fix_invalid_xml(file_path, output_file):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Escape < and > in <NAME>...</NAME>
#     content = re.sub(r'(<NAME>[^<]*?)(<)([^>]*?)(</NAME>)', r'\1&lt;\3\4', content)

#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Final XML fixed and saved: {output_file}")
#     return output_file

# fix_invalid_xml(cleaned_file, fixed_file)

# # ---------- Step 5: Convert XML to Excel ----------
# def xml_to_excel(xml_path, excel_path):
#     try:
#         tree = ET.parse(xml_path)
#         root = tree.getroot()

#         data = []
#         for ledger in root.iter("LEDGER"):
#             record = {
#                 "NAME": ledger.findtext("NAME", default=""),
#                 "PARENT": ledger.findtext("PARENT", default=""),
#                 "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
#             }
#             data.append(record)

#         if data:
#             df = pd.DataFrame(data)
#             df.to_excel(excel_path, index=False)
#             print(f"✅ Excel created: {excel_path}")
#         else:
#             print("⚠️ No LEDGER data found in the XML.")

#     except Exception as e:
#         print(f"❌ Error parsing XML: {e}")

# xml_to_excel(fixed_file, excel_file)






# import os
# import re
# import html
# import pandas as pd
# import xml.etree.ElementTree as ET

# # ---------- Step 1: User Input ----------
# base_filename = input("Enter base filename (without extension): ").strip()

# # Folder paths
# upload_dir = r"D:\xml-json-converter\uploads"
# clean_dir = r"D:\xml-json-converter\Clean_xml"
# converted_dir = r"D:\xml-json-converter\converted"

# # Create folders if not exist
# os.makedirs(upload_dir, exist_ok=True)
# os.makedirs(clean_dir, exist_ok=True)
# os.makedirs(converted_dir, exist_ok=True)

# # File paths
# raw_file = os.path.join(upload_dir, base_filename + ".xml")
# utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
# cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
# fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
# excel_file = os.path.join(converted_dir, f"{base_filename}.xlsx")

# # ---------- Step 2: Convert to UTF-8 ----------
# with open(raw_file, "r", encoding="utf-16", errors="ignore") as f:
#     content = f.read()

# with open(utf8_file, "w", encoding="utf-8") as f:
#     f.write(content)

# print(f"✅ UTF-8 conversion done: {utf8_file}")

# # ---------- Step 3: Clean XML ----------
# def clean_xml(file_path, output_file):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Remove non-printable characters
#     content = re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', content)

#     # Decode HTML entities
#     content = html.unescape(content)

#     # Fix standalone '&' not part of entities
#     content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Cleaned XML saved: {output_file}")
#     return output_file

# clean_xml(utf8_file, cleaned_file)

# # ---------- Step 4: Fix Unescaped '<' or '>' ----------
# def fix_invalid_xml(file_path, output_file):
#     with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         content = f.read()

#     # Escape < and > in <NAME>...</NAME>
#     content = re.sub(r'(<NAME>[^<]*?)(<)([^>]*?)(</NAME>)', r'\1&lt;\3\4', content)

#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(content)

#     print(f"✅ Final XML fixed and saved: {output_file}")
#     return output_file

# fix_invalid_xml(cleaned_file, fixed_file)

# # ---------- Step 5: Convert XML to Excel ----------
# def xml_to_excel(xml_path, excel_path):
#     try:
#         tree = ET.parse(xml_path)
#         root = tree.getroot()

#         data = []
#         for ledger in root.iter("LEDGER"):
#             record = {
#                 "NAME": ledger.findtext("NAME", default=""),
#                 "PARENT": ledger.findtext("PARENT", default=""),
#                 "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
#             }
#             data.append(record)

#         if data:
#             df = pd.DataFrame(data)
#             df.to_excel(excel_path, index=False)
#             print(f"✅ Excel created: {excel_path}")
#         else:
#             print("⚠️ No LEDGER data found in the XML.")

#     except Exception as e:
#         print(f"❌ Error parsing XML: {e}")

# xml_to_excel(fixed_file, excel_file)


# import os
# import re
# import html
# import pandas as pd
# import xml.etree.ElementTree as ET

# # ---------- Step 1: User Input ----------
# base_filename = input("Enter base filename (without extension): ").strip()

# # ---------- Folder paths ----------
# upload_dir = r"D:\xml-json-converter\uploads"
# clean_dir = r"D:\xml-json-converter\Clean_xml"
# converted_dir = r"D:\xml-json-converter\Converted_xlx"

# # Create folders if not exist
# os.makedirs(upload_dir, exist_ok=True)
# os.makedirs(clean_dir, exist_ok=True)
# os.makedirs(converted_dir, exist_ok=True)

# # ---------- File paths ----------
# raw_file = os.path.join(upload_dir, base_filename + ".xml")
# utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
# cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
# fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
# excel_file = os.path.join(converted_dir, f"{base_filename}.xlsx")

# # ---------- Step 2: Check for Raw File ----------
# if not os.path.exists(raw_file):
#     print(f"❌ File not found: {raw_file}")
#     exit()

# # ---------- Step 3: Skip cleaning if already cleaned ----------
# if not os.path.exists(fixed_file):
#     # ---------- Convert to UTF-8 ----------
#     with open(raw_file, "r", encoding="utf-16", errors="ignore") as f:
#         content = f.read()
#     with open(utf8_file, "w", encoding="utf-8") as f:
#         f.write(content)
#     print(f"✅ UTF-8 conversion done: {utf8_file}")

#     # ---------- Clean XML ----------
#     def clean_xml(file_path, output_file):
#         with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#             content = f.read()

#         content = re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', content)
#         content = html.unescape(content)
#         content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

#         with open(output_file, 'w', encoding='utf-8') as f:
#             f.write(content)

#         print(f"✅ Cleaned XML saved: {output_file}")
#         return output_file

#     clean_xml(utf8_file, cleaned_file)

#     # ---------- Fix Unescaped Tags ----------
#     def fix_invalid_xml(file_path, output_file):
#         with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#             content = f.read()

#         content = re.sub(r'(<NAME>[^<]*?)(<)([^>]*?)(</NAME>)', r'\1&lt;\3\4', content)

#         with open(output_file, 'w', encoding='utf-8') as f:
#             f.write(content)

#         print(f"✅ Final XML fixed and saved: {output_file}")
#         return output_file

#     fix_invalid_xml(cleaned_file, fixed_file)
# else:
#     print("⚠️ Cleaned file already exists. Skipping cleaning steps.")

# # ---------- Step 4: Convert XML to Excel ----------
# def xml_to_excel(xml_path, excel_path):
#     try:
#         tree = ET.parse(xml_path)
#         root = tree.getroot()

#         data = []
#         for ledger in root.iter("LEDGER"):
#             record = {
#                 "NAME": ledger.findtext("NAME", default=""),
#                 "PARENT": ledger.findtext("PARENT", default=""),
#                 "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
#             }
#             data.append(record)

#         if data:
#             df = pd.DataFrame(data)
#             df.to_excel(excel_path, index=False)
#             print(f"✅ Excel created: {excel_path}")
#         else:
#             print("⚠️ No LEDGER data found in the XML.")

#     except Exception as e:
#         print(f"❌ Error parsing XML: {e}")

# xml_to_excel(fixed_file, excel_file)

# # ---------- Step 5: Final Success Message ----------
# print("🎉 Process executed successfully!")



# import os
# import re
# import html
# import requests
# import pandas as pd
# import xml.etree.ElementTree as ET

# # ---------- Step 1: User Input ----------
# base_filename = input("Enter base filename (without extension): ").strip()

# # ---------- Folder paths ----------
# upload_dir = r"D:\xml-json-converter\uploads"
# clean_dir = r"D:\xml-json-converter\Clean_xml"
# converted_dir = r"D:\xml-json-converter\Converted_xlx"

# # Create folders if not exist
# os.makedirs(upload_dir, exist_ok=True)
# os.makedirs(clean_dir, exist_ok=True)
# os.makedirs(converted_dir, exist_ok=True)

# # ---------- File paths ----------
# raw_file = os.path.join(upload_dir, base_filename + ".xml")
# utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
# cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
# fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
# excel_file = os.path.join(converted_dir, f"{base_filename}.xlsx")

# # ---------- Step 2: Fetch XML from Tally and Save ----------
# if not os.path.exists(raw_file):
#     print("📡 Fetching XML from Tally...")

#     url = "http://localhost:9000"
#     xml_request = f"""
#     <ENVELOPE>
#       <HEADER>
#         <VERSION>1</VERSION>
#         <TALLYREQUEST>Export</TALLYREQUEST>
#         <TYPE>Collection</TYPE>
#         <ID>Ledger</ID>
#       </HEADER>
#       <BODY>
#         <DESC>
#           <STATICVARIABLES>
#             <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
#             <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
#           </STATICVARIABLES>
#           <TDL>
#             <TDLMESSAGE>
#               <COLLECTION NAME="Ledgers" ISMODIFY="No">
#                 <TYPE>Ledger</TYPE>
#                 <FETCH>NAME,PARENT,OPENINGBALANCE</FETCH>
#               </COLLECTION>
#             </TDLMESSAGE>
#           </TDL>
#         </DESC>
#       </BODY>
#     </ENVELOPE>
#     """

#     try:
#         response = requests.post(url, data=xml_request.encode('utf-8'))
#         response.raise_for_status()
#         with open(raw_file, 'wb') as f:
#             f.write(response.content)
#         print(f"✅ XML file saved to: {raw_file}")
#     except Exception as e:
#         print(f"❌ Failed to fetch XML from Tally: {e}")
#         exit()
# else:
#     print("📁 Raw XML file already exists. Skipping fetch.")

# # ---------- Step 3: Skip cleaning if already done ----------
# if not os.path.exists(fixed_file):
#     # Convert to UTF-8
#     with open(raw_file, "r", encoding="utf-16", errors="ignore") as f:
#         content = f.read()
#     with open(utf8_file, "w", encoding="utf-8") as f:
#         f.write(content)
#     print(f"✅ UTF-8 conversion done: {utf8_file}")

#     # Clean XML
#     def clean_xml(file_path, output_file):
#         # with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#         #     content = f.read()

#         # Try reading as UTF-8, fallback to UTF-16 if it fails
#         try:
#            with open(raw_file, "r", encoding="utf-8") as f:
#                content = f.read()
#            print("📖 Read as UTF-8.")
#         except UnicodeDecodeError:
#            with open(raw_file, "r", encoding="utf-16") as f:
#                content = f.read()
#            print("📖 Read as UTF-16.")
    

#         content = re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', content)
#         content = html.unescape(content)
#         content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

#         with open(output_file, 'w', encoding='utf-8') as f:
#             f.write(content)

#         print(f"✅ Cleaned XML saved: {output_file}")
#         return output_file

#     clean_xml(utf8_file, cleaned_file)

#     # Fix invalid XML
#     def fix_invalid_xml(file_path, output_file):
#         with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
#             content = f.read()

#         content = re.sub(r'(<NAME>[^<]*?)(<)([^>]*?)(</NAME>)', r'\1&lt;\3\4', content)

#         with open(output_file, 'w', encoding='utf-8') as f:
#             f.write(content)

#         print(f"✅ Final XML fixed and saved: {output_file}")
#         return output_file

#     fix_invalid_xml(cleaned_file, fixed_file)
# else:
#     print("⚠️ Cleaned XML already present. Skipping cleaning process.")

# # ---------- Step 4: Convert XML to Excel ----------
# def xml_to_excel(xml_path, excel_path):
#     try:
#         tree = ET.parse(xml_path)
#         root = tree.getroot()

#         data = []
#         for ledger in root.iter("LEDGER"):
#             record = {
#                 "NAME": ledger.findtext("NAME", default=""),
#                 "PARENT": ledger.findtext("PARENT", default=""),
#                 "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
#             }
#             data.append(record)

#         if data:
#             df = pd.DataFrame(data)
#             df.to_excel(excel_path, index=False)
#             print(f"✅ Excel created: {excel_path}")
#         else:
#             print("⚠️ No LEDGER data found in the XML.")

#     except Exception as e:
#         print(f"❌ Error parsing XML: {e}")

# xml_to_excel(fixed_file, excel_file)

# # ---------- Final Step ----------
# print("🎉 Whole process executed successfully!")



import os
import re
import html
import requests
import pandas as pd
import xml.etree.ElementTree as ET

# ---------- Helper: Smart Encoding Reader ----------
def read_text_smart(filepath):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        print(f"📖 Read {os.path.basename(filepath)} as UTF-8.")
        return content
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="utf-16") as f:
            content = f.read()
        print(f"📖 Read {os.path.basename(filepath)} as UTF-16.")
        return content

# ---------- Step 1: User Input ----------
base_filename = input("Enter base filename (without extension): ").strip()

# ---------- Folder paths ----------
upload_dir = r"D:\xml-json-converter\uploads"
clean_dir = r"D:\xml-json-converter\Clean_xml"
converted_dir = r"D:\xml-json-converter\Converted_xlx"

# Create folders if not exist
os.makedirs(upload_dir, exist_ok=True)
os.makedirs(clean_dir, exist_ok=True)
os.makedirs(converted_dir, exist_ok=True)

# ---------- File paths ----------
raw_file = os.path.join(upload_dir, base_filename + ".xml")
utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
excel_file = os.path.join(converted_dir, f"{base_filename}.xlsx")

# ---------- Step 2: Fetch XML from Tally and Save ----------
if not os.path.exists(raw_file):
    print("📡 Fetching XML from Tally...")

    url = "http://localhost:9000"
    xml_request = f"""
    <ENVELOPE>
      <HEADER>
        <VERSION>1</VERSION>
        <TALLYREQUEST>Export</TALLYREQUEST>
        <TYPE>Collection</TYPE>
        <ID>Ledger</ID>
      </HEADER>
      <BODY>
        <DESC>
          <STATICVARIABLES>
            <SVEXPORTFORMAT>XML</SVEXPORTFORMAT>
            <SVCURRENTCOMPANY>P.S.JIVRAJ &amp; Co.(2023-24)</SVCURRENTCOMPANY>
          </STATICVARIABLES>
          <TDL>
            <TDLMESSAGE>
              <COLLECTION NAME="Ledgers" ISMODIFY="No">
                <TYPE>Ledger</TYPE>
                <FETCH>NAME,PARENT,OPENINGBALANCE</FETCH>
              </COLLECTION>
            </TDLMESSAGE>
          </TDL>
        </DESC>
      </BODY>
    </ENVELOPE>
    """

    try:
        response = requests.post(url, data=xml_request.encode('utf-8'))
        response.raise_for_status()
        with open(raw_file, 'wb') as f:
            f.write(response.content)
        print(f"✅ XML file saved to: {raw_file}")
    except Exception as e:
        print(f"❌ Failed to fetch XML from Tally: {e}")
        exit()
else:
    print("📁 Raw XML file already exists. Skipping fetch.")

# ---------- Step 3: Clean XML ----------
if not os.path.exists(fixed_file):
    # Convert to UTF-8
    content = read_text_smart(raw_file)
    with open(utf8_file, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"✅ UTF-8 conversion done: {utf8_file}")

    # Clean XML
    def clean_xml(file_path, output_file):
        content = read_text_smart(file_path)
        content = re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', content)
        content = html.unescape(content)
        content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;', content)

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)

        print(f"✅ Cleaned XML saved: {output_file}")
        return output_file

    clean_xml(utf8_file, cleaned_file)

    # Fix invalid XML
    def fix_invalid_xml(file_path, output_file):
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()

        content = re.sub(r'(<NAME>[^<]*?)(<)([^>]*?)(</NAME>)', r'\1&lt;\3\4', content)

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)

        print(f"✅ Final XML fixed and saved: {output_file}")
        return output_file

    fix_invalid_xml(cleaned_file, fixed_file)
else:
    print("⚠️ Cleaned XML already present. Skipping cleaning process.")

# ---------- Step 4: Convert XML to Excel ----------
def xml_to_excel(xml_path, excel_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        data = []
        for ledger in root.iter("LEDGER"):
            record = {
                "NAME": ledger.findtext("NAME", default=""),
                "PARENT": ledger.findtext("PARENT", default=""),
                "OPENINGBALANCE": ledger.findtext("OPENINGBALANCE", default="")
            }
            data.append(record)

        if data:
            df = pd.DataFrame(data)
            df.to_excel(excel_path, index=False)
            print(f"✅ Excel created: {excel_path}")
        else:
            print("⚠️ No LEDGER data found in the XML.")

    except Exception as e:
        print(f"❌ Error parsing XML: {e}")

xml_to_excel(fixed_file, excel_file)

# ---------- Final Step ----------
print("🎉 Whole process executed successfully!")


