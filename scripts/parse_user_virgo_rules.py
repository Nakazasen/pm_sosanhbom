raw_text = """BOTTLE WASTE	Full_name	
FILL UP CONTAINER ASSY	Part_name	
PWB PANEL MAIN ASSY WITH SOFTWARE	Full_name	
PWB KEY-R ASSY	Full_name	
PWB KEY-L ASSY	Full_name	
PWB LCDCONT ASSY WITH SOFTWARE	Full_name	
HOLDER BELT TENSION R	Full_name	
HOLDER HEATER	Full_name	
HOLDER THERMAL-CUTOUT	Full_name	
HOLDER BELT TENSION L	Full_name	
HOLDER THERMISTOR	Full_name	
CONN.CORD ASSY S04301 HEATER CON-FUSER B	Full_name	
BUSH EXIT	Full_name	
PWB FUSER ASSY	Full_name	
PULLEY FEED P ASSY	Full_name	
ONE-WAY CLUTCH 6	Full_name	
GEAR Z23R ONEWAY	Full_name	
PWB ASSY TX	Full_name	
PWB ASSY RX WITH SOFTWARE	Full_name	
PULLEY PICKUP ASSY	Full_name	
PWB CIS CONNECT ASSY	Full_name	
CASE DLP	Full_name	
P2200 DEVE /D26	Part_name	
PWB DLP ASSY	Full_name	
PWB TPC ASSY	Full_name	
PWB MOTOR ASSY	Full_name	
FRAME CONVEYING	Full_name	
PWB ID SENSOR ASSY	Full_name	
PWB SUB ASSY	Full_name	
PWB THERMISTOR ASSY	Full_name	
PWB DRUM CONNECT ASSY	Full_name	
PWB SWITCH ASSY	Full_name	
PWB MAIN ASSY	Part_name	
PWB EXIT ASSY	Full_name	
PWB CONTAINER CONNECT ASSY	Full_name	
PWB THERMISTOR2 ASSY	Full_name	
PWB ENGINE CONNECT ASSY	Full_name	
PWB ERASER ASSY	Full_name	
PWB DRUM ASSY	Full_name	
FRAME DRUM	Full_name	
DRUM ASSY WITH PACKING	Full_name	
PWB APC A ASSY	Full_name	
HOUSING LSU	Full_name	
PWB APC B ASSY	Full_name	
P.W.BOARD ASSY PD B	Full_name	
P.W.BOARD ASSY PD A	Full_name	
PULLEY PICKUP P ASSY	Full_name	
FRAME TOP	Full_name	
ROLLER RETARD P ASSY	Full_name	
ROLLER M/P ASSY	Full_name	
PWB APC A ASSY	Part_name	
PWB APC B ASSY	Part_name	
APC ASSY B	Part_name	
PWB TPC ASSY	Part_name	
PWB DLP ASSY	Part_name	
PWB TPC ASSY	Part_name	
CASE DLP	Part_name	
PWB ERASER ASSY	Part_name	
P.W.BOARD ASSY PD B	Part_name	
HOUSING LSU	Part_name	
P.W.BOARD ASSY PD A	Part_name	
PWB APC A ASSY	Part_name"""

lines = [line.strip() for line in raw_text.strip().split("\n") if line.strip()]
print(f"Total lines parsed: {len(lines)}")
for idx, l in enumerate(lines, start=1):
    parts = l.split("\t")
    item = parts[0].strip()
    mode = parts[1].strip() if len(parts) > 1 else "Full_name"
    code = parts[2].strip() if len(parts) > 2 and parts[2].strip() else None
    print(f"{idx:02d}: item='{item}' | mode='{mode}' | code={code}")
