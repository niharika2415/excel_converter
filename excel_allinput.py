import pandas as pd
from io import BytesIO, StringIO
import json
from typing import Any

def universal_to_excel(data: Any, file_name="exported_data.xlsx") -> BytesIO:
        """
    Convert various input formats into Excel (in memory).

    Supported formats:
    - dict (single row or dict-of-lists)
    - list/tuple (flat row or table)
    - list of dicts (records)
    - pandas DataFrame
    - CSV string
    - Markdown table string
    - JSON string
    """
    
        df= None
    
        # Case 1: Already a dataframe
        if isinstance(data, pd.DataFrame):
            df= data
            
        # Case 2: List of dicts (records)
        elif isinstance(data, list) and all(isinstance(i,dict) for i in data):
            df= pd.DataFrame(data)
        
        # Case 3: Dict (single row or multiple row)
        elif isinstance(data, dict):
            if all(isinstance(v, (list, tuple)) for v in data.values()):       #multi-row dict
                df= pd.DataFrame(data)
            else:                                                  
                df= pd.DataFrame(data)        #single-row dict
        
        # Case 4: List of lists/tuples (table with header row)
        elif isinstance(data, (list, tuple)) and all(isinstance(i, (list, tuple)) for i in data):
            if len(data) > 1:
                df= pd.DataFrame(data[1:], column=data[0])
            else:
                df= pd.DataFrame(data)
        
        # Case 5: Flat list/tuple (single row)
        elif isinstance(data, (list, tuple)):
            df= pd.DataFrame([data])
            
        # Case 6: CSV string
        elif isinstance(data, str) and "," in data:
            df= pd.read_csv(StringIO(data))
            
        # Case 7: Markdown Table String
        elif isinstance(data, str) and  "|" in data:
            try:
                df= pd.read_csv(StringIO(data), sep= "|", engine="python")
                df= df.dropna(axis=1, how= "all") #remove empty cols
                df.columns= [c.strip(" -") for c in df.columns]
                df= df.applymap(lambda x: str(x).strip(" -") if isinstance(x, str) else x)
            except Exception:
                raise ValueError("Invalid markdown table format")
            
        # Case 8: JSON string
        elif isinstance(data, str):
            try:
                parsed= json.loads(data)
                return universal_to_excel(parsed, file_name)
            except json.JSONDecodeError:
                raise ValueError("Unsupported string format")
            
        else:
            raise ValueError("Unsupported Input Format")
        
        # Export DataFrame → Excel (in memory)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        output.seek(0)
        return output