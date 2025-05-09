import pandas as pd
import os

def split_and_filter_csv():
    input_csv = r"C:\Users\joshu\OneDrive - NSWGOV\Rental Data Model\Calculated rental data\Median rents\output data\suburb_rent_data.csv"
    output_dir = r"C:\Users\joshu\OneDrive - NSWGOV\Rental Data Model\Calculated rental data\Median rents\output data\split_files"

    os.makedirs(output_dir, exist_ok=True)

    # Read the full CSV
    try:
        df = pd.read_csv(input_csv, parse_dates=["month"])
    except Exception as e:
        print(f"Failed to read CSV: {e}")
        return

    # Add a year column
    df["year"] = df["month"].dt.year

    # Create one file per year
    for year in df["year"].unique():
        year_df = df[df["year"] == year]
        year_file = os.path.join(output_dir, f"rent_data_{year}.csv")
        year_df.to_csv(year_file, index=False)
        print(f"Wrote file for {year}: {year_file}")

    # Filter dwellings only
    dwellings_df = df[df["property_type"].str.lower() == "dwellings"]
    dwellings_file = os.path.join(output_dir, "rent_data_dwellings_only.csv")
    dwellings_df.to_csv(dwellings_file, index=False)
    print(f"Wrote filtered dwellings file: {dwellings_file}")

if __name__ == "__main__":
    split_and_filter_csv()
