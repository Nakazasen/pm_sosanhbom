import os
import glob
import xml.etree.ElementTree as ET

appdata = os.environ.get('APPDATA', '')
localappdata = os.environ.get('LOCALAPPDATA', '')

sap_paths = [
    os.path.join(appdata, 'SAP', 'Common'),
    os.path.join(appdata, 'SAP', 'SAP GUI'),
    os.path.join(localappdata, 'SAP', 'Common'),
]

print("=== Checking SAP Landscape / INI files ===")
for base in sap_paths:
    if os.path.exists(base):
        print(f"\nScanning: {base}")
        for f in os.listdir(base):
            full_path = os.path.join(base, f)
            print(f"  File: {f} (size={os.path.getsize(full_path)})")
            if f.endswith('.xml'):
                try:
                    tree = ET.parse(full_path)
                    root = tree.getroot()
                    # Find services/connections
                    for elem in root.iter():
                        if 'name' in elem.attrib:
                            name = elem.attrib['name']
                            if 'P1J' in name or 'ERP' in name or 'VN' in name:
                                print(f"    Found connection node: tag={elem.tag}, attrib={elem.attrib}")
                except Exception as e:
                    print(f"    Error parsing {f}: {e}")
            elif f.endswith('.ini'):
                try:
                    with open(full_path, 'r', errors='ignore') as fp:
                        for line in fp:
                            if 'P1J' in line or 'ERP' in line or 'VN' in line:
                                print(f"    INI line: {line.strip()}")
                except Exception as e:
                    print(f"    Error reading {f}: {e}")
