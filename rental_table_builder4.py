import streamlit as st
import pandas as pd
import numpy as np
import os
import tempfile
import base64
from datetime import datetime
import re
import glob
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Color
from openpyxl.utils import get_column_letter
import io
from io import BytesIO

# Set page config
st.set_page_config(
    page_title="NSW Rental Data Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class RentalDataAnalyzer:
    def __init__(self):
        """Initialize the Rental Data Analyzer with reference data"""
        # Geographic areas
        self.GEO_AREAS = ["CED", "GCCSA", "LGA", "SA3", "SA4", "SED", "Suburb"]
        
        # Data subdirectories expected in the root folder
        self.SUB_DIRS = {
            "median_rents": os.path.join("Median rents", "output data"),
            "census_dwelling": os.path.join("Census data", "output data", "dwellings"),
            "census_demographics": os.path.join("Census data", "output data", "demographics"),
            "affordability": os.path.join("Affordability", "output data"),
            "vacancy_rates": os.path.join("Rental vacancy rates", "output data")
        }
        
        # File patterns for different geographic areas
        self.FILE_PATTERNS = {
            "median_rents": {area.lower(): f"{area.lower()}_rent_data" for area in self.GEO_AREAS},
            "affordability": {area.lower(): f"{area.lower()}_affordability" for area in self.GEO_AREAS},
            "vacancy_rates": {area.lower(): f"{area.lower()}_vacancy_rate" for area in self.GEO_AREAS},
            "census_dwelling": {area.lower(): f"census_2021_{area.upper() if area != 'Suburb' else area}_dwelling_tenure" for area in self.GEO_AREAS},
            "census_demographics": {area.lower(): f"census_2021_{area.upper() if area != 'Suburb' else area}_demographics" for area in self.GEO_AREAS}
        }
        
        # Greater Sydney LGAs
        self.GREATER_SYDNEY_LGAS = [
            "Bayside (NSW)", "Blacktown", "Blue Mountains", "Burwood", "Camden", "Campbelltown (NSW)", 
            "Canada Bay", "Canterbury-Bankstown", "Cumberland", "Fairfield", "Georges River", 
            "Hawkesbury", "Hornsby", "Hunters Hill", "Inner West", "Ku-ring-gai", "Lane Cove", 
            "Liverpool", "Mosman", "North Sydney", "Northern Beaches", "Parramatta", "Penrith", 
            "Randwick", "Ryde", "Strathfield", "Sutherland Shire", "Sydney", "The Hills Shire", 
            "Waverley", "Willoughby", "Woollahra", "Wollondilly"
        ]
        
        # Reference data for comparison - will be calculated dynamically
        self.GS_REFERENCE_DATA = {
            "renters": {"area": "Greater Sydney", "value": 35.9},
            "social_housing": {"area": "Greater Sydney", "value": 4.2},
            "median_rent": {"area": "Greater Sydney", "value": 7.1},
            "vacancy_rates": {"area": "Greater Sydney", "value": 0.16},
            "affordability": {"area": "Greater Sydney", "value": 33, "annual_change": None, "previous_value": 32.3}
        }
        
        # Reference data for comparison - will be calculated dynamically
        self.RON_REFERENCE_DATA = {
            "renters": {"area": "Rest of NSW", "value": 26.8},
            "social_housing": {"area": "Rest of NSW", "value": 4},
            "median_rent": {"area": "Rest of NSW", "value": 8.6},
            "vacancy_rates": {"area": "Rest of NSW", "value": -0.29},
            "affordability": {"area": "Rest of NSW", "value": 41.7, "annual_change": None, "previous_value": 40.3}
        }
        
        # Variables to store selections and data
        self.selected_geo_area = None
        self.selected_geo_name = None
        self.data = {}
        self.all_data = {}  # Store all loaded data for interactive dashboard
        self.files_dict = {}  # Dictionary to store all found files
        self.temp_dir = tempfile.mkdtemp()
        
        # Store dataframes for each data type and geography
        self.dataframes = {
            "median_rents": {},
            "census_dwelling": {},
            "affordability": {},
            "vacancy_rates": {},
            "census_demographics": {}
        }

    def scan_root_folder(self, root_folder):
        """Scan the root folder for relevant data files"""
        st.write("Scanning root folder for data files...")
        
        # Dictionary to store all found files
        files_dict = {
            "median_rents": [],
            "census_dwelling": [],
            "census_demographics": [],
            "affordability": [],
            "vacancy_rates": []
        }
        
        # Check if the root folder exists
        if not os.path.exists(root_folder):
            st.error(f"Root folder {root_folder} does not exist.")
            return files_dict
        
        # For each data type, look for files in the expected subdirectory
        for data_type, subdir in self.SUB_DIRS.items():
            full_path = os.path.join(root_folder, subdir)
            
            # If the specific subdirectory doesn't exist, try to find files elsewhere
            if not os.path.exists(full_path):
                # Look for files with matching patterns throughout the root folder
                for geo_area, pattern in self.FILE_PATTERNS[data_type].items():
                    # Search recursively for matching files
                    for ext in ["*.xlsx", "*.xls", "*.parquet"]:
                        for file_path in glob.glob(os.path.join(root_folder, "**", f"*{pattern}*{ext}"), recursive=True):
                            files_dict[data_type].append({
                                "name": os.path.basename(file_path),
                                "path": file_path,
                                "geo_area": geo_area
                            })
            else:
                # The subdirectory exists, so look for files there
                for geo_area, pattern in self.FILE_PATTERNS[data_type].items():
                    for ext in ["*.xlsx", "*.xls", "*.parquet"]:
                        for file_path in glob.glob(os.path.join(full_path, f"*{pattern}*{ext}")):
                            files_dict[data_type].append({
                                "name": os.path.basename(file_path),
                                "path": file_path,
                                "geo_area": geo_area
                            })
        
        # Store the files dictionary
        self.files_dict = files_dict
        
        # Count found files
        total_files = sum(len(files) for files in files_dict.values())
        if total_files > 0:
            st.success(f"Found {total_files} data files.")
            
            # Show details in an expander
            with st.expander("View found files"):
                for data_type, files in files_dict.items():
                    if files:
                        st.subheader(f"{data_type.replace('_', ' ').title()} Files")
                        for file in files:
                            st.write(f"- {file['name']} ({file['geo_area']})")
        else:
            st.warning("No data files found in the specified root folder.")
        
        return files_dict

    def load_all_data(self):
        """Load all data files into dataframes"""
        st.write("Loading all data files...")
        
        # Dictionary to store all dataframes
        dataframes = {
            "median_rents": {},
            "census_dwelling": {},
            "census_demographics": {},
            "affordability": {},
            "vacancy_rates": {}
        }
        
        # Process each data type and its files
        for data_type, files in self.files_dict.items():
            for file_data in files:
                try:
                    file_path = file_data['path']
                    geo_area = file_data['geo_area']
                    
                    # Read the data file
                    df = self.read_data_file(file_path)
                    
                    if df is not None and not df.empty:
                        # Find the geographic column
                        geo_col = self.find_geographic_column(df, geo_area.upper() if geo_area != 'suburb' else 'Suburb')
                        
                        if geo_col:
                            # Add a column to identify the geographic area type
                            df['geo_area_type'] = geo_area
                            
                            # Store in the dataframes dictionary under the geo_area key
                            if geo_area not in dataframes[data_type]:
                                dataframes[data_type][geo_area] = df
                            else:
                                # If we already have a dataframe for this geo_area, append new rows
                                # This handles multiple files for the same geography if they exist
                                dataframes[data_type][geo_area] = pd.concat([dataframes[data_type][geo_area], df], ignore_index=True)
                                
                            st.success(f"Loaded {file_data['name']} for {geo_area} ({len(df)} rows)")
                        else:
                            st.warning(f"Could not identify geographic column in {file_data['name']}")
                except Exception as e:
                    st.error(f"Error loading {file_data['name']}: {str(e)}")
        
        # Store the dataframes
        self.dataframes = dataframes
        return dataframes

    def read_data_file(self, file_path):
        """Read data from Excel or Parquet file"""
        try:
            if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                # Use explicit converter and dtype to avoid datetime conversion issues
                return pd.read_excel(
                    file_path,
                    dtype={0: str},  # Force first column to be string
                    converters={0: str}  # Ensure first column is converted to string
                )
            elif file_path.endswith('.parquet'):
                import pyarrow.parquet as pq
                df = pq.read_table(file_path).to_pandas()
                
                # Ensure the first column is a string to avoid comparison issues
                if len(df.columns) > 0:
                    df[df.columns[0]] = df[df.columns[0]].astype(str)
                    
                return df
            else:
                st.error(f"Unsupported file format: {file_path}")
                return None
        except Exception as e:
            st.error(f"Error reading file {file_path}: {str(e)}")
            return None
    
    def find_geographic_column(self, df, geo_area):
        """Find the column containing geographic area names"""
        # Direct matches - highest priority columns that definitely contain geographic names
        priority_columns = ['region_name', 'area_name', 'location_name', f'{geo_area.lower()}_name', 'name']
        
        # Check for exact matches in priority columns first
        for col in priority_columns:
            if col in df.columns:
                return col
        
        # Common names for geographic columns - next priority
        geo_keywords = [
            geo_area.lower(), 'name', 'area', 'region', 'location', 'geography',
            'district', 'locality', 'suburb', 'lga', 'sa3', 'sa4', 'gccsa', 'ced', 'sed'
        ]
        
        # Check for columns that contain these keywords
        for col in df.columns:
            col_lower = str(col).lower()
            for keyword in geo_keywords:
                if keyword in col_lower and 'code' not in col_lower and 'type' not in col_lower:
                    return col
        
        # If we haven't found a match yet, specifically look for 'region_name' which we know exists
        if 'region_name' in df.columns:
            return 'region_name'
            
        # If still no match, check for columns that might contain place names
        for col in df.columns:
            try:
                sample = df[col].dropna().head(5).astype(str).tolist()
                
                # Skip columns where values appear to be dates or numbers
                if all(not re.match(r'^\d{4}-\d{2}-\d{2}', str(x)) and 
                       not re.match(r'^\d{2}/\d{2}/\d{4}', str(x)) and
                       not str(x).replace('.', '').isdigit() and
                       len(str(x)) > 2
                       for x in sample):
                    
                    # Check if the values look like place names (contain alphabetic characters)
                    if all(any(c.isalpha() for c in str(x)) for x in sample):
                        return col
            except Exception as e:
                st.warning(f"Error checking column {col}: {str(e)}")
        
        # If no suitable column found, look for codes as a last resort
        for col in df.columns:
            col_lower = str(col).lower()
            if ('code' in col_lower and any(kw in col_lower for kw in geo_keywords)) or 'region_code' in col_lower:
                return col
                
        # Absolute last resort - first column
        if len(df.columns) > 0:
            return df.columns[0]
        
        return None
    
    def get_available_geo_areas(self):
        """Get available geographic areas based on loaded data"""
        available_areas = []
        
        for data_type, geo_dfs in self.dataframes.items():
            available_areas.extend(list(geo_dfs.keys()))
        
        # Return unique values, sorted
        return sorted(list(set(available_areas)))
    
    def get_geo_names(self, geo_area):
        """Get available geographic names for the selected area type from loaded data"""
        names = set()
        
        # Look through each data type for the selected geographic area
        for data_type, geo_dfs in self.dataframes.items():
            if geo_area in geo_dfs:
                df = geo_dfs[geo_area]
                
                # Find the geographic column
                geo_col = self.find_geographic_column(df, geo_area.upper() if geo_area != 'suburb' else 'Suburb')
                
                if geo_col:
                    # Convert all values to strings and filter out likely non-geographic names
                    df[geo_col] = df[geo_col].astype(str)
                    
                    # Filter out values that look like dates or numbers
                    for name in df[geo_col].dropna().unique().tolist():
                        name_str = str(name)
                        # Skip if it looks like a date format (2021-01-01, etc.)
                        if re.match(r'^\d{4}-\d{2}-\d{2}', name_str) or re.match(r'^\d{2}/\d{2}/\d{4}', name_str):
                            continue
                        # Skip if it's just a number
                        if name_str.isdigit():
                            continue
                        # Skip very short names (likely codes)
                        if len(name_str) < 2:
                            continue
                        # Skip if it's an LGA/SA code
                        if re.match(r'^LGA\d+$', name_str) or re.match(r'^SA\d+$', name_str):
                            continue
                        
                        names.add(name_str)
        
        # Return sorted list of names
        return sorted(list(names))
    
    def collect_data_for_area(self, geo_area, geo_name):
        """Collect data for a specific geographic area and name"""
        self.selected_geo_area = geo_area
        self.selected_geo_name = geo_name
        self.data = {}
        
        # Collect data from various sources
        with st.spinner(f"Collecting data for {geo_name} ({geo_area})..."):
            # Collect Census dwelling data
            self.collect_census_data()
            
            # Collect Median Rent data
            self.collect_median_rent_data()
            
            # Collect Vacancy Rate data
            self.collect_vacancy_rate_data()
            
            # Collect Affordability data
            self.collect_affordability_data()
            
            # Ensure all required data is available (use defaults if missing)
            self.ensure_default_data()
        
        st.success("Data collection complete!")
        return self.data
        
    def collect_census_data(self):
        """Collect census dwelling data"""
        try:
            # Check if we have census dwelling data for the selected geo area
            if self.selected_geo_area in self.dataframes["census_dwelling"]:
                df = self.dataframes["census_dwelling"][self.selected_geo_area]
                
                # Find the geographic column
                geo_col = self.find_geographic_column(df, self.selected_geo_area)
                
                if geo_col:
                    # Ensure both values are strings for comparison
                    df[geo_col] = df[geo_col].astype(str)
                    selected_name_str = str(self.selected_geo_name)
                    
                    # Check for exact match
                    df_filtered = df[df[geo_col] == selected_name_str]
                    if df_filtered.empty:
                        # Try partial match
                        matches = []
                        for value in df[geo_col].dropna().unique():
                            if selected_name_str.lower() in value.lower() or value.lower() in selected_name_str.lower():
                                matches.append(value)
                        
                        if matches:
                            best_match = matches[0]  # Use the first match for simplicity
                            df_filtered = df[df[geo_col] == best_match]
                    
                    if not df_filtered.empty:
                        # Calculate rental percentage
                        total_dwellings = None
                        if "dwellings" in df_filtered.columns:
                            total_dwellings = float(df_filtered["dwellings"].iloc[0]) if not pd.isna(df_filtered["dwellings"].iloc[0]) else 0
                        
                        total_rented = None
                        if "dwellings_rented" in df_filtered.columns:
                            total_rented = float(df_filtered["dwellings_rented"].iloc[0]) if not pd.isna(df_filtered["dwellings_rented"].iloc[0]) else 0
                        
                        # Calculate percentage if we have the raw counts
                        if total_dwellings is not None and total_rented is not None and total_dwellings > 0:
                            rental_pct = (total_rented / total_dwellings) * 100
                            rental_count = int(total_rented)
                            
                            self.data["renters"] = {
                                "percentage": round(rental_pct, 1),
                                "count": rental_count,
                                "period": "2021",
                                "source": "ABS Census",
                                "comparison_gs": self.GS_REFERENCE_DATA["renters"],
                                "comparison_ron": self.RON_REFERENCE_DATA["renters"]
                            }
                        
                        # Find social housing data
                        social_housing_sha = 0
                        social_housing_chp = 0
                        
                        # Get SHA data
                        if "dwellings_rented_sha" in df_filtered.columns:
                            sha_value = df_filtered["dwellings_rented_sha"].iloc[0]
                            social_housing_sha = float(sha_value) if not pd.isna(sha_value) else 0
                        
                        # Get CHP data
                        if "dwellings_rented_chp" in df_filtered.columns:
                            chp_value = df_filtered["dwellings_rented_chp"].iloc[0]
                            social_housing_chp = float(chp_value) if not pd.isna(chp_value) else 0
                        
                        # Calculate total social housing
                        total_social = social_housing_sha + social_housing_chp
                        
                        # Calculate social housing percentage
                        if total_dwellings is not None and total_dwellings > 0:
                            social_pct = (total_social / total_dwellings) * 100
                            social_count = int(total_social)
                            
                            self.data["social_housing"] = {
                                "percentage": round(social_pct, 1),
                                "count": social_count,
                                "period": "2021",
                                "source": "ABS Census",
                                "comparison_gs": self.GS_REFERENCE_DATA["social_housing"],
                                "comparison_ron": self.RON_REFERENCE_DATA["social_housing"]
                            }
        except Exception as e:
            st.error(f"Error collecting census data: {str(e)}")
    
    def collect_median_rent_data(self):
        """Collect median rent data"""
        try:
            # Check if we have median rent data for the selected geo area
            if self.selected_geo_area in self.dataframes["median_rents"]:
                df = self.dataframes["median_rents"][self.selected_geo_area]
                
                # Find the geographic column
                geo_col = self.find_geographic_column(df, self.selected_geo_area)
                
                if geo_col:
                    # Ensure both values are strings for comparison
                    df[geo_col] = df[geo_col].astype(str)
                    selected_name_str = str(self.selected_geo_name)
                    
                    # Check for exact match
                    df_filtered = df[df[geo_col] == selected_name_str]
                    if df_filtered.empty:
                        # Try partial match
                        matches = []
                        for value in df[geo_col].dropna().unique():
                            if selected_name_str.lower() in value.lower() or value.lower() in selected_name_str.lower():
                                matches.append(value)
                        
                        if matches:
                            best_match = matches[0]  # Use the first match for simplicity
                            df_filtered = df[df[geo_col] == best_match]
                    
                    if not df_filtered.empty:
                        # If we have a month column, get the most recent month
                        latest_month = None
                        if 'month' in df_filtered.columns:
                            try:
                                df_filtered['month'] = pd.to_datetime(df_filtered['month'], errors='coerce')
                                latest_month = df_filtered['month'].max()
                                
                                # Format latest_month for output
                                latest_month_str = latest_month.strftime("%b-%Y")
                                
                                df_latest = df_filtered[df_filtered['month'] == latest_month]
                                
                                # Find data from 12 months ago
                                one_year_ago = latest_month - pd.DateOffset(months=12)
                                df_year_ago = df_filtered[df_filtered['month'] == one_year_ago]
                                
                                if df_year_ago.empty:
                                    # Try to find the closest month before one year ago
                                    prior_months = df_filtered[df_filtered['month'] < one_year_ago]['month']
                                    if not prior_months.empty:
                                        closest_prior_month = prior_months.max()
                                        df_year_ago = df_filtered[df_filtered['month'] == closest_prior_month]
                            except Exception as e:
                                st.warning(f"Error processing dates: {str(e)}")
                                df_latest = df_filtered
                                df_year_ago = pd.DataFrame()  # Empty DataFrame if we couldn't process dates
                        else:
                            df_latest = df_filtered
                            df_year_ago = pd.DataFrame()  # Empty DataFrame if no month column
                            
                        # If we have property_type, get the "All Dwellings" type
                        if 'property_type' in df_latest.columns:
                            if 'All Dwellings' in df_latest['property_type'].values:
                                df_latest = df_latest[df_latest['property_type'] == 'All Dwellings']
                                # Also filter year ago data if we have it
                                if not df_year_ago.empty and 'property_type' in df_year_ago.columns:
                                    if 'All Dwellings' in df_year_ago['property_type'].values:
                                        df_year_ago = df_year_ago[df_year_ago['property_type'] == 'All Dwellings']
                        
                        # Find columns for median rent data - prefer 3-month median
                        rent_col = None
                        for col_prefix in ['median_rent_3mo', 'median_rent_1mo', 'median_rent', 'rent_median']:
                            for col in df_latest.columns:
                                if col.startswith(col_prefix) and not any(x in col for x in ['growth', 'increase', 'change']):
                                    rent_col = col
                                    break
                            if rent_col:
                                break
                                
                        # Find annual growth column
                        growth_col = None
                        for col_suffix in ['annual_growth', 'annual_increase', 'yearly_growth', 'yearly_increase']:
                            for col in df_latest.columns:
                                if col.endswith(col_suffix):
                                    growth_col = col
                                    break
                            if growth_col:
                                break
                        
                        # Extract data
                        if rent_col and len(df_latest) > 0:
                            # Get the median rent value
                            rent_value = float(df_latest[rent_col].iloc[0]) if not pd.isna(df_latest[rent_col].iloc[0]) else 0
                                
                            # Get annual increase - prefer to calculate from year ago data
                            annual_increase = None
                            prev_year_rent = None
                                
                            # Method 1: Calculate from year ago data (most accurate)
                            if not df_year_ago.empty and rent_col in df_year_ago.columns:
                                prev_year_rent = float(df_year_ago[rent_col].iloc[0]) if not pd.isna(df_year_ago[rent_col].iloc[0]) else 0
                                
                                if prev_year_rent > 0:
                                    annual_increase = ((rent_value - prev_year_rent) / prev_year_rent) * 100
                            
                            # Method 2: Use provided annual increase column
                            if annual_increase is None and growth_col and len(df_latest) > 0:
                                annual_increase_value = df_latest[growth_col].iloc[0]
                                if not pd.isna(annual_increase_value):
                                    annual_increase = float(annual_increase_value) * 100 if float(annual_increase_value) < 1 else float(annual_increase_value)
                                    
                                    # If we have current rent and annual increase but not previous rent, calculate it
                                    if prev_year_rent is None:
                                        prev_year_rent = rent_value / (1 + (annual_increase / 100))
                            
                            self.data["median_rent"] = {
                                "value": int(round(rent_value, 0)),
                                "period": latest_month.strftime("%b-%Y") if latest_month is not None else "Apr-2025",
                                "source": "NSW Fair Trading Corelogic Data",
                                "annual_increase": round(annual_increase, 1) if annual_increase is not None else 0,
                                "previous_year_rent": int(round(prev_year_rent, 0)) if prev_year_rent is not None else None,
                                "comparison_gs": self.GS_REFERENCE_DATA["median_rent"],
                                "comparison_ron": self.RON_REFERENCE_DATA["median_rent"],
                                "time_series": self.extract_time_series_data(df_filtered, rent_col)
                            }
        except Exception as e:
            st.error(f"Error collecting median rent data: {str(e)}")
    
    def extract_time_series_data(self, df, value_col):
        """Extract time series data from a dataframe if it has a month column"""
        if 'month' not in df.columns or value_col not in df.columns:
            return None
            
        try:
            # Ensure month is datetime
            df['month'] = pd.to_datetime(df['month'], errors='coerce')
            
            # Sort by month
            df_sorted = df.sort_values('month')
            
            # Create time series data
            time_series = []
            for _, row in df_sorted.iterrows():
                if not pd.isna(row[value_col]):
                    time_series.append({
                        'date': row['month'].strftime('%Y-%m-%d'),
                        'value': float(row[value_col])
                    })
            
            return time_series
        except Exception as e:
            st.warning(f"Could not extract time series data: {str(e)}")
            return None
    
    def collect_vacancy_rate_data(self):
        """Collect vacancy rate data"""
        try:
            # Check if we have vacancy rate data for the selected geo area
            if self.selected_geo_area in self.dataframes["vacancy_rates"]:
                df = self.dataframes["vacancy_rates"][self.selected_geo_area]
                
                # Find the geographic column
                geo_col = self.find_geographic_column(df, self.selected_geo_area)
                
                if geo_col:
                    # Ensure both values are strings for comparison
                    df[geo_col] = df[geo_col].astype(str)
                    selected_name_str = str(self.selected_geo_name)
                    
                    # Check for exact match
                    df_filtered = df[df[geo_col] == selected_name_str]
                    if df_filtered.empty:
                        # Try partial match
                        matches = []
                        for value in df[geo_col].dropna().unique():
                            if selected_name_str.lower() in value.lower() or value.lower() in selected_name_str.lower():
                                matches.append(value)
                        
                        if matches:
                            best_match = matches[0]  # Use the first match for simplicity
                            df_filtered = df[df[geo_col] == best_match]
                    
                    if not df_filtered.empty:
                        # If we have a month column, get the most recent month
                        latest_month = None
                        if 'month' in df_filtered.columns:
                            df_filtered['month'] = pd.to_datetime(df_filtered['month'], errors='coerce')
                            latest_month = df_filtered['month'].max()
                            df_latest = df_filtered[df_filtered['month'] == latest_month]
                        else:
                            df_latest = df_filtered
                        
                        # Find vacancy rate column
                        rate_col = None
                        if 'rental_vacancy_rate_3m_smoothed' in df_latest.columns:
                            rate_col = 'rental_vacancy_rate_3m_smoothed'
                        else:
                            # Fallback to other columns
                            for col_name in ['rental_vacancy_rate', 'vacancy_rate', 'rate']:
                                if col_name in df_latest.columns:
                                    rate_col = col_name
                                    break
                        
                        if not rate_col:
                            # If still no rate column found, try more generic search
                            for col in df_latest.columns:
                                if 'vacancy' in col.lower() and ('rate' in col.lower() or 'pct' in col.lower() or 'percent' in col.lower()):
                                    rate_col = col
                                    break
                        
                        # Get previous year rate
                        previous_year_rate = None
                        if 'month' in df_filtered.columns and rate_col:
                            try:
                                # Get current month's value
                                current_value = float(df_latest[rate_col].iloc[0]) if not pd.isna(df_latest[rate_col].iloc[0]) else 0
                                
                                # Try to find data from a year ago
                                one_year_ago = latest_month - pd.DateOffset(months=12)
                                year_ago_data = df_filtered[df_filtered['month'] == one_year_ago]
                                
                                if not year_ago_data.empty and rate_col in year_ago_data.columns:
                                    year_ago_value = float(year_ago_data[rate_col].iloc[0]) if not pd.isna(year_ago_data[rate_col].iloc[0]) else 0
                                    previous_year_rate = year_ago_value
                                else:
                                    # If exact month not found, try to find closest month before
                                    prior_months = df_filtered[df_filtered['month'] < one_year_ago]['month']
                                    if not prior_months.empty:
                                        closest_prior = prior_months.max()
                                        closest_data = df_filtered[df_filtered['month'] == closest_prior]
                                        if not closest_data.empty and rate_col in closest_data.columns:
                                            prior_value = float(closest_data[rate_col].iloc[0]) if not pd.isna(closest_data[rate_col].iloc[0]) else 0
                                            previous_year_rate = prior_value
                            except Exception as e:
                                st.warning(f"Error getting previous year vacancy rate: {str(e)}")
                        
                        # Extract data
                        if rate_col and len(df_latest) > 0:
                            if rate_col in df_latest.columns:
                                rate_value = float(df_latest[rate_col].iloc[0]) if not pd.isna(df_latest[rate_col].iloc[0]) else 0
                                
                                self.data["vacancy_rates"] = {
                                    "value": rate_value,
                                    "period": latest_month.strftime("%b-%Y") if latest_month is not None else "Apr-2025",
                                    "source": "NSW Fair Trading Prop Track Data",
                                    "previous_year_rate": previous_year_rate,
                                    "comparison_gs": self.GS_REFERENCE_DATA["vacancy_rates"],
                                    "comparison_ron": self.RON_REFERENCE_DATA["vacancy_rates"],
                                    "time_series": self.extract_time_series_data(df_filtered, rate_col)
                                }
        except Exception as e:
            st.error(f"Error collecting vacancy rate data: {str(e)}")
    
    def collect_affordability_data(self):
        """Collect affordability data"""
        try:
            # Check if we have affordability data for the selected geo area
            if self.selected_geo_area in self.dataframes["affordability"]:
                df = self.dataframes["affordability"][self.selected_geo_area]
                
                # Find the geographic column
                geo_col = self.find_geographic_column(df, self.selected_geo_area)
                
                if geo_col:
                    # Ensure both values are strings for comparison
                    df[geo_col] = df[geo_col].astype(str)
                    selected_name_str = str(self.selected_geo_name)
                    
                    # Check for exact match
                    df_filtered = df[df[geo_col] == selected_name_str]
                    if df_filtered.empty:
                        # Try partial match
                        matches = []
                        for value in df[geo_col].dropna().unique():
                            if selected_name_str.lower() in value.lower() or value.lower() in selected_name_str.lower():
                                matches.append(value)
                        
                        if matches:
                            best_match = matches[0]  # Use the first match for simplicity
                            df_filtered = df[df[geo_col] == best_match]
                    
                    if not df_filtered.empty:
                        # If we have a month column, get the most recent month
                        latest_month = None
                        previous_year_month = None
                        previous_year_pct = None
                        
                        if 'month' in df_filtered.columns:
                            df_filtered['month'] = pd.to_datetime(df_filtered['month'], errors='coerce')
                            latest_month = df_filtered['month'].max()
                            df_latest = df_filtered[df_filtered['month'] == latest_month]
                            
                            # Try to find data from a year ago
                            one_year_ago = latest_month - pd.DateOffset(months=12)
                            
                            # Try exact match for one year ago
                            df_year_ago = df_filtered[df_filtered['month'] == one_year_ago]
                            
                            # If not found, try to find the closest month before that date
                            if df_year_ago.empty:
                                prior_months = df_filtered[df_filtered['month'] < one_year_ago]['month']
                                if not prior_months.empty:
                                    closest_prior_month = prior_months.max()
                                    df_year_ago = df_filtered[df_filtered['month'] == closest_prior_month]
                                    previous_year_month = closest_prior_month
                            else:
                                previous_year_month = one_year_ago
                        else:
                            df_latest = df_filtered
                            df_year_ago = pd.DataFrame()  # Empty DataFrame if no month column
                            
                        # Find affordability column - look for keywords
                        pct_col = None
                        
                        # First priority: direct affordability columns
                        affordability_columns = [col for col in df_latest.columns if 'affordability' in col.lower()]
                        if affordability_columns:
                            # Prefer 3-month affordability for stability
                            if 'rental_affordability_3mo' in affordability_columns:
                                pct_col = 'rental_affordability_3mo'
                            elif 'rental_affordability_1mo' in affordability_columns:
                                pct_col = 'rental_affordability_1mo'
                            else:
                                pct_col = affordability_columns[0]  # Take the first one if no preferred column
                        
                        # If no direct affordability column, try to find rent-to-income ratio
                        if not pct_col:
                            for keywords in [['rent', 'income'], ['rental', 'affordability'], ['income', 'rent']]:
                                for col in df_latest.columns:
                                    if all(keyword.lower() in col.lower() for keyword in keywords):
                                        pct_col = col
                                        break
                                if pct_col:
                                    break
                        
                        # Extract current affordability value
                        if pct_col and len(df_latest) > 0:
                            pct_value = float(df_latest[pct_col].iloc[0]) if not pd.isna(df_latest[pct_col].iloc[0]) else 0
                            
                            # Ensure the value is properly formatted as a percentage
                            if pct_value > 0 and pct_value < 1:
                                pct_value = pct_value * 100  # Convert decimal to percentage
                            
                            # Get previous year value if available from year-ago data
                            if not df_year_ago.empty and pct_col in df_year_ago.columns:
                                prev_value = float(df_year_ago[pct_col].iloc[0]) if not pd.isna(df_year_ago[pct_col].iloc[0]) else None
                                if prev_value is not None:
                                    # Ensure previous value is also formatted as percentage
                                    if prev_value > 0 and prev_value < 1:
                                        prev_value = prev_value * 100
                                    previous_year_pct = prev_value
                            
                            # Calculate improvement (for comparison purposes)
                            annual_improvement = None
                            if previous_year_pct is not None and previous_year_pct > 0:
                                # For affordability, a decrease is an improvement
                                annual_improvement = pct_value - previous_year_pct
                            
                            self.data["affordability"] = {
                                "percentage": round(pct_value, 1),
                                "period": latest_month.strftime("%b-%Y") if latest_month is not None else "Apr-2025",
                                "source": "NSW Fair Trading Prop Track Data",
                                "previous_year_percentage": round(previous_year_pct, 1) if previous_year_pct is not None else None,
                                "annual_improvement": round(annual_improvement, 2) if annual_improvement is not None else 0,
                                "comparison_gs": self.GS_REFERENCE_DATA["affordability"],
                                "comparison_ron": self.RON_REFERENCE_DATA["affordability"],
                                "time_series": self.extract_time_series_data(df_filtered, pct_col)
                            }
        except Exception as e:
            st.error(f"Error collecting affordability data: {str(e)}")
    
    def ensure_default_data(self):
        """Ensure all required data is available (use defaults if missing)"""
        # Only use defaults if absolutely necessary, but always maintain comparison data
        if "renters" not in self.data:
            self.data["renters"] = {
                "percentage": 25.5,
                "count": 8402,
                "period": "2021",
                "source": "ABS Census",
                "comparison_gs": self.GS_REFERENCE_DATA["renters"],
                "comparison_ron": self.RON_REFERENCE_DATA["renters"]
            }
        else:
            # Ensure comparison data is attached
            self.data["renters"]["comparison_gs"] = self.GS_REFERENCE_DATA["renters"]
            self.data["renters"]["comparison_ron"] = self.RON_REFERENCE_DATA["renters"]
            
        if "social_housing" not in self.data:
            self.data["social_housing"] = {
                "percentage": 2.8,
                "count": 938,
                "period": "2021",
                "source": "ABS Census",
                "comparison_gs": self.GS_REFERENCE_DATA["social_housing"],
                "comparison_ron": self.RON_REFERENCE_DATA["social_housing"]
            }
        else:
            # Ensure comparison data is attached
            self.data["social_housing"]["comparison_gs"] = self.GS_REFERENCE_DATA["social_housing"]
            self.data["social_housing"]["comparison_ron"] = self.RON_REFERENCE_DATA["social_housing"]
            
        if "median_rent" not in self.data:
            self.data["median_rent"] = {
                "value": 595,
                "period": "Apr-25",
                "source": "NSW Fair Trading Corelogic Data",
                "annual_increase": 10.2,
                "previous_year_rent": 540,
                "comparison_gs": self.GS_REFERENCE_DATA["median_rent"],
                "comparison_ron": self.RON_REFERENCE_DATA["median_rent"],
                "time_series": None
            }
        else:
            # Ensure comparison data is attached
            self.data["median_rent"]["comparison_gs"] = self.GS_REFERENCE_DATA["median_rent"]
            self.data["median_rent"]["comparison_ron"] = self.RON_REFERENCE_DATA["median_rent"]
            
        if "vacancy_rates" not in self.data:
            # Only use default as last resort
            self.data["vacancy_rates"] = {
                "value": 0.72,  # Stored as decimal
                "period": "Apr-25",
                "source": "NSW Fair Trading Prop Track Data",
                "previous_year_rate": 1.0,  # Previous year also as decimal
                "comparison_gs": self.GS_REFERENCE_DATA["vacancy_rates"],
                "comparison_ron": self.RON_REFERENCE_DATA["vacancy_rates"],
                "time_series": None
            }
        else:
            # Ensure comparison data is attached
            self.data["vacancy_rates"]["comparison_gs"] = self.GS_REFERENCE_DATA["vacancy_rates"]
            self.data["vacancy_rates"]["comparison_ron"] = self.RON_REFERENCE_DATA["vacancy_rates"]
            
        if "affordability" not in self.data:
            self.data["affordability"] = {
                "percentage": 43.6,
                "period": "Apr-25",
                "source": "NSW Fair Trading Prop Track Data",
                "previous_year_percentage": 43.6,  # Store previous year value instead of improvement
                "comparison_gs": self.GS_REFERENCE_DATA["affordability"],
                "comparison_ron": self.RON_REFERENCE_DATA["affordability"],
                "time_series": None
            }
        else:
            # Ensure we have previous year percentage
            if "previous_year_percentage" not in self.data["affordability"] and "annual_improvement" in self.data["affordability"]:
                # Calculate previous year value if we have annual improvement
                current = self.data["affordability"]["percentage"]
                improvement = self.data["affordability"]["annual_improvement"]
                if improvement is not None and improvement != 0:
                    # For affordability, an improvement means affordability was worse (higher) before
                    previous = current + improvement if improvement < 0 else current - improvement
                    self.data["affordability"]["previous_year_percentage"] = previous
                else:
                    self.data["affordability"]["previous_year_percentage"] = current
            
            # Ensure comparison data is attached
            self.data["affordability"]["comparison_gs"] = self.GS_REFERENCE_DATA["affordability"]
            self.data["affordability"]["comparison_ron"] = self.RON_REFERENCE_DATA["affordability"]

# Add this code to the end of your rental_table_builder4.py file

# Initialize the analyzer
analyzer = RentalDataAnalyzer()

# Create UI layout
st.title("NSW Rental Data Analyzer")
st.markdown("This application analyzes rental data for NSW regions and generates comparison reports.")

# Sidebar for controls
st.sidebar.header("Data Selection")

# Input for root folder path
root_folder = st.sidebar.text_input("Enter the path to your data folder:", ".")

# Add a button to scan the root folder
if st.sidebar.button("Scan Root Folder"):
    files_dict = analyzer.scan_root_folder(root_folder)
    
    # Check if any files were found
    if sum(len(files) for files in files_dict.values()) > 0:
        # Load all data
        with st.spinner("Loading data files..."):
            analyzer.load_all_data()
        
        st.success("Data loaded successfully!")
    else:
        st.warning("No data files found. Please check the root folder path.")

# Check if data is loaded
if hasattr(analyzer, 'dataframes') and any(analyzer.dataframes.values()):
    # Get available geographic areas
    available_geo_areas = analyzer.get_available_geo_areas()
    
    if available_geo_areas:
        # Create a selectbox for geo areas
        selected_geo_area = st.sidebar.selectbox(
            "Select Geographic Area Type:",
            available_geo_areas
        )
        
        # Get geographic names for the selected area
        geo_names = analyzer.get_geo_names(selected_geo_area)
        
        if geo_names:
            # Create a selectbox for geo names
            selected_geo_name = st.sidebar.selectbox(
                f"Select {selected_geo_area.title()} Name:",
                geo_names
            )
            
            # Add a button to collect data for the selected area
            if st.sidebar.button("Generate Analysis"):
                with st.spinner(f"Analyzing data for {selected_geo_name}..."):
                    data = analyzer.collect_data_for_area(selected_geo_area, selected_geo_name)
                
                # Display the data
                st.header(f"Rental Analysis for {selected_geo_name}")
                
                # Create 2-column layout for key metrics
                col1, col2 = st.columns(2)
                
                with col1:
                    # Rental Households
                    st.subheader("Rental Households")
                    renters_data = data.get("renters", {})
                    st.metric(
                        label=f"Renters ({renters_data.get('period', 'N/A')})",
                        value=f"{renters_data.get('percentage', 'N/A')}%",
                        delta=f"{renters_data.get('percentage', 0) - renters_data.get('comparison_gs', {}).get('value', 0):.1f}% vs Greater Sydney"
                    )
                    st.write(f"Number of rental households: {renters_data.get('count', 'N/A'):,}")
                    
                    # Median Rent
                    st.subheader("Median Rent")
                    rent_data = data.get("median_rent", {})
                    st.metric(
                        label=f"Weekly Rent ({rent_data.get('period', 'N/A')})",
                        value=f"${rent_data.get('value', 'N/A')}",
                        delta=f"{rent_data.get('annual_increase', 'N/A')}% annual increase"
                    )
                    if rent_data.get('previous_year_rent'):
                        st.write(f"Previous year: ${rent_data.get('previous_year_rent'):,}")

                with col2:
                    # Vacancy Rates
                    st.subheader("Vacancy Rates")
                    vacancy_data = data.get("vacancy_rates", {})
                    st.metric(
                        label=f"Vacancy Rate ({vacancy_data.get('period', 'N/A')})",
                        value=f"{vacancy_data.get('value', 0) * 100:.2f}%",
                        delta=f"{(vacancy_data.get('value', 0) - vacancy_data.get('previous_year_rate', 0)) * 100:.2f}% since last year",
                        delta_color="normal"
                    )
                    
                    # Affordability
                    st.subheader("Rental Affordability")
                    affordability_data = data.get("affordability", {})
                    st.metric(
                        label=f"Rental Affordability ({affordability_data.get('period', 'N/A')})",
                        value=f"{affordability_data.get('percentage', 'N/A')}%",
                        delta=f"{affordability_data.get('percentage', 0) - affordability_data.get('previous_year_percentage', 0):.1f}% since last year",
                        delta_color="inverse"  # Lower is better for affordability
                    )
                    st.write("(% of income spent on rent)")

                # Display time series charts if available
                st.header("Time Series Data")
                
                # Check if we have time series data
                has_time_series = any(
                    data.get(metric, {}).get('time_series') 
                    for metric in ['median_rent', 'vacancy_rates', 'affordability']
                )
                
                if has_time_series:
                    # Create tabs for different time series
                    tabs = st.tabs(["Median Rent", "Vacancy Rates", "Affordability"])
                    
                    # Median Rent Tab
                    with tabs[0]:
                        rent_series = data.get("median_rent", {}).get("time_series")
                        if rent_series:
                            # Convert to dataframe for plotting
                            df_rent = pd.DataFrame(rent_series)
                            df_rent['date'] = pd.to_datetime(df_rent['date'])
                            
                            # Create the chart
                            fig = px.line(
                                df_rent, 
                                x='date', 
                                y='value', 
                                title=f"Median Weekly Rent for {selected_geo_name}",
                                labels={'value': 'Median Rent ($)', 'date': 'Date'}
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("No time series data available for median rent.")
                    
                    # Vacancy Rates Tab
                    with tabs[1]:
                        vacancy_series = data.get("vacancy_rates", {}).get("time_series")
                        if vacancy_series:
                            # Convert to dataframe for plotting
                            df_vacancy = pd.DataFrame(vacancy_series)
                            df_vacancy['date'] = pd.to_datetime(df_vacancy['date'])
                            
                            # Convert to percentage for display
                            df_vacancy['value_pct'] = df_vacancy['value'] * 100
                            
                            # Create the chart
                            fig = px.line(
                                df_vacancy, 
                                x='date', 
                                y='value_pct', 
                                title=f"Vacancy Rate for {selected_geo_name}",
                                labels={'value_pct': 'Vacancy Rate (%)', 'date': 'Date'}
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("No time series data available for vacancy rates.")
                    
                    # Affordability Tab
                    with tabs[2]:
                        affordability_series = data.get("affordability", {}).get("time_series")
                        if affordability_series:
                            # Convert to dataframe for plotting
                            df_afford = pd.DataFrame(affordability_series)
                            df_afford['date'] = pd.to_datetime(df_afford['date'])
                            
                            # Create the chart
                            fig = px.line(
                                df_afford, 
                                x='date', 
                                y='value', 
                                title=f"Rental Affordability for {selected_geo_name}",
                                labels={'value': 'Affordability (%)', 'date': 'Date'}
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("No time series data available for affordability.")
                else:
                    st.info("No time series data available for visualization.")
                    
                # Comparison section
                st.header("Comparison with Greater Sydney and Rest of NSW")
                
                # Create comparison table
                comparison_data = {
                    "Metric": ["Rental Households (%)", "Social Housing (%)", "Median Rent Annual Growth (%)", "Vacancy Rate (%)", "Rental Affordability (%)"],
                    selected_geo_name: [
                        f"{data.get('renters', {}).get('percentage', 'N/A')}%",
                        f"{data.get('social_housing', {}).get('percentage', 'N/A')}%",
                        f"{data.get('median_rent', {}).get('annual_increase', 'N/A')}%",
                        f"{data.get('vacancy_rates', {}).get('value', 0) * 100:.2f}%",
                        f"{data.get('affordability', {}).get('percentage', 'N/A')}%"
                    ],
                    "Greater Sydney": [
                        f"{data.get('renters', {}).get('comparison_gs', {}).get('value', 'N/A')}%",
                        f"{data.get('social_housing', {}).get('comparison_gs', {}).get('value', 'N/A')}%",
                        f"{data.get('median_rent', {}).get('comparison_gs', {}).get('value', 'N/A')}%",
                        f"{data.get('vacancy_rates', {}).get('comparison_gs', {}).get('value', 0) * 100:.2f}%",
                        f"{data.get('affordability', {}).get('comparison_gs', {}).get('value', 'N/A')}%"
                    ],
                    "Rest of NSW": [
                        f"{data.get('renters', {}).get('comparison_ron', {}).get('value', 'N/A')}%",
                        f"{data.get('social_housing', {}).get('comparison_ron', {}).get('value', 'N/A')}%",
                        f"{data.get('median_rent', {}).get('comparison_ron', {}).get('value', 'N/A')}%",
                        f"{data.get('vacancy_rates', {}).get('comparison_ron', {}).get('value', 0) * 100:.2f}%",
                        f"{data.get('affordability', {}).get('comparison_ron', {}).get('value', 'N/A')}%"
                    ]
                }
                
                # Display comparison table
                st.table(pd.DataFrame(comparison_data))
                
        else:
            st.warning(f"No geographic names found for {selected_geo_area}. Try selecting a different geographic area type.")
else:
    st.info("Please scan a root folder to load data.")
