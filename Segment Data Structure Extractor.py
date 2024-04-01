import pandas as pd

def excel_to_custom_dict(file_path, sheet_name):
    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Convert the DataFrame to a list of dictionaries with the desired structure
        data_list = []
        for _, row in df.iterrows():
            data_list.append({
                "Period": row["Period"],
                'Demand Units': row['Demand Units'],
                "Sales Units": row["Sales Units"],
                "Segment % of Total Industry": row["Seg % of Total Industry"],
                "Expected Growth Rate": row["Expected Growth Rate"],
                "Desired Feature": row["Desired Feature"],
                "Desired Portability": row["Desired Portability"],
                "Desired Age": row["Desired Age"],
                "Desired Style Factor": row["Desired Style Factor"],
                "Desired Price": row["Desired Price"],
                "Importance of Positioning": row["Importance of Positioning"],
                "Importance of Age": row["Importance of Age"],
                "Importance of Price": row["Importance of Price"],
                "Importance of Style": row["Importance of Style"],
                "Segment": row["Segment"],
            })

        return data_list

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# Example usage
file_path = 'C:/Users/Admin/Desktop/SM Marketing/figures/MatPlot Data.xlsx'
# Replace with the actual path to your Excel file
sheet_name = ('MatLab Data Segments')
# Replace with the actual sheet name in your Excel file

result_list = excel_to_custom_dict(file_path, sheet_name)

if result_list:
    print("Excel data converted to custom dictionary format:")
    for i, item in enumerate(result_list):
        print(item, end=", " if i < len(result_list) - 1 else "")
    print()
else:
    print("Failed to convert Excel data to custom dictionary format.")
