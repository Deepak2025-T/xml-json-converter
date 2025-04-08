import requests
import os
from converter import json_to_excel, xml_to_excel

def fetch_file_from_url(url, output_folder):
    response = requests.get(url)
    file_type = 'json' if url.endswith('.json') else 'xml'
    filename = os.path.basename(url)
    temp_path = os.path.join('uploads', filename)

    with open(temp_path, 'wb') as f:
        f.write(response.content)

    if file_type == 'json':
        out_path = json_to_excel.convert_json_to_excel(temp_path, output_folder)
    else:
        out_path = xml_to_excel.convert_xml_to_excel(temp_path, output_folder)

    return out_path, file_type, filename