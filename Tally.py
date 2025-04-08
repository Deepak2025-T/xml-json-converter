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

# ---------- Updated: Clean XML ----------
# ---------- Updated: Clean XML with Better Handling ----------


def clean_xml(file_path, output_file):
    content = read_text_smart(file_path)

    # Remove raw control characters (ASCII 0x00–0x1F except tab, LF, CR)
    content = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', content)

    # Remove invalid decimal char references like &#4;
    content = re.sub(r'&#(?:[0-8]|1[0-9]|2[0-9]|3[01]);', '', content)

    # Remove invalid hex char references like &#x1F;
    content = re.sub(r'&#x(?:0?[0-8bcef]|1[0-9a-fA-F]|2[0-9aB]);', '', content)

    # Fix LEDGER NAME attributes (escaping &, ", etc.)
    def fix_ledger_tag(match):
        name_attr = match.group(1)
        rest = match.group(2)
        name_cleaned = html.escape(name_attr, quote=True)
        return f'<LEDGER NAME="{name_cleaned}"{rest}>'

    content = re.sub(r'<LEDGER\s+NAME="(.*?)"([^>]*)>', fix_ledger_tag, content)

    # Escape & in <NAME> and <NAME.LIST> text content
    content = re.sub(r'(<NAME>)([^<]*?)(</NAME>)', lambda m: f"{m.group(1)}{html.escape(m.group(2))}{m.group(3)}", content)
    content = re.sub(r'(<NAME\.LIST[^>]*>)([^<]*?)(</NAME\.LIST>)', lambda m: f"{m.group(1)}{html.escape(m.group(2))}{m.group(3)}", content)

    # Clean PARENT tag content too
    content = re.sub(r'(<PARENT[^>]*>)([^<]*?)(</PARENT>)', lambda m: f"{m.group(1)}{html.escape(m.group(2))}{m.group(3)}", content)

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"✅ Cleaned XML saved: {output_file}")
    return output_file


# ---------- Updated: Fix Invalid XML Safely ----------
def fix_invalid_xml(file_path, output_file):
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()

    # Escape < and > inside tags that could still break parsing
    content = re.sub(
        r'(<[A-Z.]+>)([^<]*?)(</[A-Z.]+>)',
        lambda m: f"{m.group(1)}{html.escape(m.group(2))}{m.group(3)}",
        content
    )

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"✅ Final XML fixed and saved: {output_file}")
    return output_file


# ---------- Step 1: User Input ----------
base_filename = input("Enter base filename (without extension): ").strip()

# ---------- Folder paths ----------
upload_dir = r"D:\xml-json-converter\uploads"
clean_dir = r"D:\xml-json-converter\Clean_xml"
converted_dir = r"D:\xml-json-converter\Converted_xlx"

os.makedirs(upload_dir, exist_ok=True)
os.makedirs(clean_dir, exist_ok=True)
os.makedirs(converted_dir, exist_ok=True)

# ---------- File paths ----------
raw_file = os.path.join(upload_dir, base_filename + ".xml")
utf8_file = os.path.join(clean_dir, f"{base_filename}_utf8.xml")
cleaned_file = os.path.join(clean_dir, f"{base_filename}_cleaned.xml")
fixed_file = os.path.join(clean_dir, f"{base_filename}_final.xml")
excel_file = os.path.join(converted_dir, f"{base_filename}.xlsx")

# ---------- Step 2: Fetch XML from Tally ----------
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
    content = read_text_smart(raw_file)
    with open(utf8_file, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"✅ UTF-8 conversion done: {utf8_file}")

    clean_xml(utf8_file, cleaned_file)
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
