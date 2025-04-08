import pandas as pd
import xml.etree.ElementTree as ET
import os

def convert_xml_to_excel(filepath, output_folder):
    tree = ET.parse(filepath)
    root = tree.getroot()
    all_records = []
    for child in root:
        record = {elem.tag: elem.text for elem in child}
        all_records.append(record)
    df = pd.DataFrame(all_records)
    filename = os.path.splitext(os.path.basename(filepath))[0] + '.xlsx'
    out_path = os.path.join(output_folder, filename)
    df.to_excel(out_path, index=False)
    return out_path
