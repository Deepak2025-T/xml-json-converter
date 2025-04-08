import pandas as pd
import os

def convert_json_to_excel(filepath, output_folder):
    df = pd.read_json(filepath)
    filename = os.path.splitext(os.path.basename(filepath))[0] + '.xlsx'
    out_path = os.path.join(output_folder, filename)
    df.to_excel(out_path, index=False)
    return out_path
