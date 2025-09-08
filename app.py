import streamlit as st 
from excel_allinput import universal_to_excel

st.title("Data To Excel")

# Text Area For User Input
user_input= st.text_area("Enter your data:")

# Convert Button
if st.button("Convert To Excel"):
    try:
        # Try conversion
        excel_file= universal_to_excel(user_input, file_name="converted.xlsx")
        
        # Provide download link 
        st.download_button(
            label= "Download Excel File",
            data= excel_file,
            file_name= "converted.xlsx",
            mime= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("Conversion Successful")
    
    except Exception as e:
        st.error(f"Error: {str(e)}")