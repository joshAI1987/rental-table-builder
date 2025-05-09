import pandas as pd
import os

def convert_parquet_to_csv():
    parquet_file = r"C:\Users\joshu\OneDrive - NSWGOV\Rental Data Model\Calculated rental data\Median rents\output data\suburb_rent_data.parquet"
    output_file = r"C:\Users\joshu\OneDrive - NSWGOV\Rental Data Model\Calculated rental data\Median rents\output data\suburb_rent_data.csv"

    if not os.path.isfile(parquet_file):
        print(f"Error: File '{parquet_file}' does not exist.")
        return

    try:
        df = pd.read_parquet(parquet_file)
    except Exception as e:
        print(f"Failed to read Parquet file: {e}")
        return

    try:
        df.to_csv(output_file, index=False)
        print(f"Successfully converted to: {output_file}")
    except Exception as e:
        print(f"Failed to write CSV file: {e}")

if __name__ == "__main__":
    convert_parquet_to_csv()
