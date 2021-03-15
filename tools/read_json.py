import os
from config import base_path
import json

file_name = base_path + os.sep + "data" + os.sep + '11.json'
with open(file_name, 'r', encoding='utf-8') as f:
    datas = json.load(f)

# print(datas)

