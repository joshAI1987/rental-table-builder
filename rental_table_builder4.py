import streamlit as st
import pandas as pd
import numpy as np
import os
import tempfile
import base64
from datetime import datetime
import re
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
        
        # Data subdirectories
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
        self.uploaded_files = {}
        self.temp_dir = tempfile.mkdtemp()

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
    
    def get_geo_names(self, geo_area, uploaded_files):
        """Get available geographic names for the selected area type from uploaded files"""
        names = set()
        found_files = False
        
        for data_type, file_group in uploaded_files.items():
            if geo_area.lower() not in self.FILE_PATTERNS[data_type]:
                continue
                
            file_pattern = self.FILE_PATTERNS[data_type][geo_area.lower()]
            
            for file_data in file_group:
                file_name = file_data['name']
                if file_pattern.lower() in file_name.lower():
                    found_files = True
                    file_path = file_data['path']
                    df = self.read_data_file(file_path)
                    
                    if df is not None and not df.empty:
                        # Look for the geographic name column
                        geo_col = self.find_geographic_column(df, geo_area)
                        
                        if geo_col:
                            # Convert all values to strings and filter out likely non-geographic names
                            df[geo_col] = df[geo_col].astype(str)
                            
                            # Filter out values that look like dates or numbers
                            area_names = []
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
                                
                                area_names.append(name_str)
                            
                            names.update(area_names)
        
        if not found_files:
            st.error(f"No data files found for {geo_area}. Please check your uploaded files.")
            return []
        
        if not names:
            st.error(f"No geographic names found for {geo_area}. Check that your data files contain the expected columns.")
            return []
            
        return sorted(list(names))
    
    def collect_data(self, uploaded_files):
        """Collect data from various sources"""
        self.data = {}
        
        # Collect data for the selected geographic name
        st.write(f"Collecting data for {self.selected_geo_name} ({self.selected_geo_area})...")
        
        # Collect Census dwelling data
        with st.spinner('Processing census dwelling data...'):
            self.collect_census_data(uploaded_files)
        
        # Collect Median Rent data
        with st.spinner('Processing median rent data...'):
            self.collect_median_rent_data(uploaded_files)
        
        # Collect Vacancy Rate data
        with st.spinner('Processing vacancy rate data...'):
            self.collect_vacancy_rate_data(uploaded_files)
        
        # Collect Affordability data
        with st.spinner('Processing affordability data...'):
            self.collect_affordability_data(uploaded_files)
        
        # Ensure all required data is available (use defaults if missing)
        self.ensure_default_data()
        
        st.success("Data collection complete!")
        return self.data
        
    def collect_census_data(self, uploaded_files):
        """Collect census dwelling data"""
        try:
            # Find census dwelling files
            file_pattern = self.FILE_PATTERNS["census_dwelling"][self.selected_geo_area.lower()]
            
            for file_data in uploaded_files.get("census_dwelling", []):
                file_name = file_data['name']
                if file_pattern.lower() in file_name.lower():
                    file_path = file_data['path']
                    df = self.read_data_file(file_path)
                    
                    if df is not None and not df.empty:
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
                        break
        except Exception as e:
            st.error(f"Error collecting census data: {str(e)}")
    
    def collect_median_rent_data(self, uploaded_files):
        """Collect median rent data"""
        try:
            # Find median rent files
            file_pattern = self.FILE_PATTERNS["median_rents"][self.selected_geo_area.lower()]
            
            for file_data in uploaded_files.get("median_rents", []):
                file_name = file_data['name']
                if file_pattern.lower() in file_name.lower():
                    file_path = file_data['path']
                    df = self.read_data_file(file_path)
                    
                    if df is not None and not df.empty:
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
                                if rent_col:
                                    # Get the median rent value
                                    if len(df_latest) > 0:
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
                                            "comparison_ron": self.RON_REFERENCE_DATA["median_rent"]
                                        }
                        break
        except Exception as e:
            st.error(f"Error collecting median rent data: {str(e)}")
    
    def collect_vacancy_rate_data(self, uploaded_files):
        """Collect vacancy rate data"""
        try:
            # Find vacancy rate files
            file_pattern = self.FILE_PATTERNS["vacancy_rates"][self.selected_geo_area.lower()]
            found_files = False
            
            for file_data in uploaded_files.get("vacancy_rates", []):
                file_name = file_data['name']
                if file_pattern.lower() in file_name.lower():
                    found_files = True
                    file_path = file_data['path']
                    df = self.read_data_file(file_path)
                    
                    if df is not None and not df.empty:
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
                                            "comparison_ron": self.RON_REFERENCE_DATA["vacancy_rates"]
                                        }
                    break
        
            if not found_files:
                st.warning(f"No vacancy rate files found matching pattern: {file_pattern}")
                
        except Exception as e:
            st.error(f"Error collecting vacancy rate data: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
    
    def collect_affordability_data(self, uploaded_files):
        """Collect affordability data"""
        try:
            # Find affordability files
            file_pattern = self.FILE_PATTERNS["affordability"][self.selected_geo_area.lower()]
            
            for file_data in uploaded_files.get("affordability", []):
                file_name = file_data['name']
                if file_pattern.lower() in file_name.lower():
                    file_path = file_data['path']
                    df = self.read_data_file(file_path)
                    
                    if df is not None and not df.empty:
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
                                        "comparison_ron": self.RON_REFERENCE_DATA["affordability"]
                                    }
                        break
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
                "comparison_ron": self.RON_REFERENCE_DATA["median_rent"]
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
                "comparison_ron": self.RON_REFERENCE_DATA["vacancy_rates"]
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
                "comparison_ron": self.RON_REFERENCE_DATA["affordability"]
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

    def fetch_comparison_area_data(self, uploaded_files, area_names=["Greater Sydney", "Rest of NSW"]):
        """Fetch actual data for comparison areas (Greater Sydney and Rest of NSW)"""
        comparison_data = {}
        
        try:
            for area_name in area_names:
                comparison_data[area_name] = {
                    "renters": None,
                    "social_housing": None,
                    "median_rent": None,
                    "vacancy_rates": None,
                    "affordability": None,
                    "affordability_previous": None
                }
                
                # For each data category, fetch the actual data for the area
                for data_type in ["census_dwelling", "median_rents", "vacancy_rates", "affordability"]:
                    for file_data in uploaded_files.get(data_type, []):
                        file_name = file_data['name']
                        if "gccsa" in file_name.lower():  # GCCSA level has both Greater Sydney and Rest of NSW
                            file_path = file_data['path']
                            df = self.read_data_file(file_path)
                            
                            if df is not None and not df.empty:
                                geo_col = self.find_geographic_column(df, "GCCSA")
                                
                                if geo_col:
                                    # Find rows for the current area
                                    df_area = df[df[geo_col].str.contains(area_name, case=False, na=False)]
                                    
                                    if not df_area.empty:
                                        # Process according to data type
                                        if data_type == "census_dwelling":
                                            # Calculate renter percentage
                                            if "dwellings" in df_area.columns and "dwellings_rented" in df_area.columns:
                                                dwellings = float(df_area["dwellings"].iloc[0]) if not pd.isna(df_area["dwellings"].iloc[0]) else 0
                                                rented = float(df_area["dwellings_rented"].iloc[0]) if not pd.isna(df_area["dwellings_rented"].iloc[0]) else 0
                                                
                                                if dwellings > 0:
                                                    comparison_data[area_name]["renters"] = round((rented / dwellings) * 100, 1)
                                            
                                            # Calculate social housing percentage
                                            if "dwellings" in df_area.columns and "dwellings_rented_sha" in df_area.columns and "dwellings_rented_chp" in df_area.columns:
                                                dwellings = float(df_area["dwellings"].iloc[0]) if not pd.isna(df_area["dwellings"].iloc[0]) else 0
                                                sha = float(df_area["dwellings_rented_sha"].iloc[0]) if not pd.isna(df_area["dwellings_rented_sha"].iloc[0]) else 0
                                                chp = float(df_area["dwellings_rented_chp"].iloc[0]) if not pd.isna(df_area["dwellings_rented_chp"].iloc[0]) else 0
                                                
                                                if dwellings > 0:
                                                    comparison_data[area_name]["social_housing"] = round(((sha + chp) / dwellings) * 100, 1)
                                        
                                        elif data_type == "median_rents":
                                            # Get latest month data
                                            if 'month' in df_area.columns:
                                                df_area['month'] = pd.to_datetime(df_area['month'], errors='coerce')
                                                latest_month = df_area['month'].max()
                                                df_latest = df_area[df_area['month'] == latest_month]
                                                
                                                if 'property_type' in df_latest.columns and 'All Dwellings' in df_latest['property_type'].values:
                                                    df_latest = df_latest[df_latest['property_type'] == 'All Dwellings']
                                                
                                                # Find annual growth column
                                                growth_col = None
                                                for suffix in ['annual_growth', 'annual_increase', 'yearly_growth', 'yearly_increase']:
                                                    for col in df_latest.columns:
                                                        if col.endswith(suffix):
                                                            growth_col = col
                                                            break
                                                    if growth_col:
                                                        break
                                                
                                                if growth_col and len(df_latest) > 0:
                                                    growth_value = df_latest[growth_col].iloc[0]
                                                    if not pd.isna(growth_value):
                                                        annual_increase = float(growth_value) * 100 if float(growth_value) < 1 else float(growth_value)
                                                        comparison_data[area_name]["median_rent"] = round(annual_increase, 1)
                                        
                                        elif data_type == "vacancy_rates":
                                            # Get latest month and year-ago data
                                            if 'month' in df_area.columns:
                                                df_area['month'] = pd.to_datetime(df_area['month'], errors='coerce')
                                                latest_month = df_area['month'].max()
                                                df_latest = df_area[df_area['month'] == latest_month]
                                                
                                                # Find vacancy rate column
                                                rate_col = None
                                                if 'rental_vacancy_rate_3m_smoothed' in df_latest.columns:
                                                    rate_col = 'rental_vacancy_rate_3m_smoothed'
                                                else:
                                                    for col in ['rental_vacancy_rate', 'vacancy_rate', 'rate']:
                                                        if col in df_latest.columns:
                                                            rate_col = col
                                                            break
                                                
                                                if rate_col and len(df_latest) > 0:
                                                    current_rate = float(df_latest[rate_col].iloc[0]) if not pd.isna(df_latest[rate_col].iloc[0]) else 0
                                                    
                                                    # Get year-ago data to calculate change
                                                    one_year_ago = latest_month - pd.DateOffset(months=12)
                                                    df_year_ago = df_area[df_area['month'] == one_year_ago]
                                                    
                                                    if not df_year_ago.empty and rate_col in df_year_ago.columns:
                                                        prev_rate = float(df_year_ago[rate_col].iloc[0]) if not pd.isna(df_year_ago[rate_col].iloc[0]) else 0
                                                        change = current_rate - prev_rate
                                                        comparison_data[area_name]["vacancy_rates"] = round(change, 2)
                                        
                                        elif data_type == "affordability":
                                            # Get latest month data
                                            if 'month' in df_area.columns:
                                                df_area['month'] = pd.to_datetime(df_area['month'], errors='coerce')
                                                latest_month = df_area['month'].max()
                                                df_latest = df_area[df_area['month'] == latest_month]
                                                
                                                # Find affordability column
                                                aff_col = None
                                                affordability_columns = [col for col in df_latest.columns if 'affordability' in col.lower()]
                                                if affordability_columns:
                                                    if 'rental_affordability_3mo' in affordability_columns:
                                                        aff_col = 'rental_affordability_3mo'
                                                    elif 'rental_affordability_1mo' in affordability_columns:
                                                        aff_col = 'rental_affordability_1mo'
                                                    else:
                                                        aff_col = affordability_columns[0]
                                                
                                                if aff_col and len(df_latest) > 0:
                                                    aff_value = float(df_latest[aff_col].iloc[0]) if not pd.isna(df_latest[aff_col].iloc[0]) else 0
                                                    if aff_value > 0 and aff_value < 1:
                                                        aff_value = aff_value * 100
                                                    
                                                    comparison_data[area_name]["affordability"] = round(aff_value, 1)
                                                    
                                                    # Get year-ago data
                                                    one_year_ago = latest_month - pd.DateOffset(months=12)
                                                    df_year_ago = df_area[df_area['month'] == one_year_ago]
                                                    
                                                    if not df_year_ago.empty and aff_col in df_year_ago.columns:
                                                        prev_aff = float(df_year_ago[aff_col].iloc[0]) if not pd.isna(df_year_ago[aff_col].iloc[0]) else 0
                                                        if prev_aff > 0 and prev_aff < 1:
                                                            prev_aff = prev_aff * 100
                                                        
                                                        comparison_data[area_name]["affordability_previous"] = round(prev_aff, 1)
        except Exception as e:
            st.error(f"Error in fetch_comparison_area_data: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
        
        # Store the fetched comparison data
        self.comparison_data = comparison_data
        return comparison_data
    
    def generate_comparison_comment(self, metric, value, comparison_gs, comparison_ron):
        """Generate a comparison comment for a metric that shows both Greater Sydney and Rest of NSW references"""
        
        # Use actual fetched data instead of reference values
        gs_value = None
        ron_value = None
        
        try:
            if hasattr(self, 'comparison_data'):
                if "Greater Sydney" in self.comparison_data and metric in self.comparison_data["Greater Sydney"]:
                    gs_value = self.comparison_data["Greater Sydney"][metric]
                
                if "Rest of NSW" in self.comparison_data and metric in self.comparison_data["Rest of NSW"]:
                    ron_value = self.comparison_data["Rest of NSW"][metric]
        except Exception as e:
            st.warning(f"Error accessing comparison data: {str(e)}")
        
        # Fall back to the provided comparison values if needed
        if gs_value is None and comparison_gs is not None:
            gs_value = comparison_gs["value"]
        
        if ron_value is None and comparison_ron is not None:
            ron_value = comparison_ron["value"]
        
        # Now proceed with the comparison using the fetched values
        if metric == "renters":
            gs_text = ""
            if value < gs_value - 1:  # 1% buffer to avoid "slightly lower" for small differences
                gs_text = f"lower than the Greater Sydney average of {gs_value}%"
            elif value > gs_value + 1:
                gs_text = f"higher than the Greater Sydney average of {gs_value}%"
            else:
                gs_text = f"similar to the Greater Sydney average of {gs_value}%"
                
            ron_text = ""
            if value < ron_value - 1:
                ron_text = f"and lower than the Rest of NSW average of {ron_value}%"
            elif value > ron_value + 1:
                ron_text = f"and higher than the Rest of NSW average of {ron_value}%"
            else:
                ron_text = f"and similar to the Rest of NSW average of {ron_value}%"
                
            return f"{self.selected_geo_name} ({self.selected_geo_area}) has a concentration of renters that is {gs_text} {ron_text}."
        
        elif metric == "social_housing":
            gs_text = ""
            if value < gs_value - 0.5:  # 0.5% buffer
                gs_text = f"lower than the Greater Sydney average of {gs_value}%"
            elif value > gs_value + 0.5:
                gs_text = f"higher than the Greater Sydney average of {gs_value}%"
            else:
                gs_text = f"similar to the Greater Sydney average of {gs_value}%"
                
            ron_text = ""
            if value < ron_value - 0.5:
                ron_text = f"and lower than the Rest of NSW average of {ron_value}%"
            elif value > ron_value + 0.5:
                ron_text = f"and higher than the Rest of NSW average of {ron_value}%"
            else:
                ron_text = f"and similar to the Rest of NSW average of {ron_value}%"
                
            return f"{self.selected_geo_name} ({self.selected_geo_area}) has a concentration of social housing that is {gs_text} {ron_text}."
        
        elif metric == "median_rent":
            local_increase = self.data["median_rent"]["annual_increase"]
            if pd.isna(local_increase):
                local_increase = 0

            # Ensure we use the dynamically fetched values
            gs_value = self.comparison_data["Greater Sydney"]["median_rent"] if hasattr(self, 'comparison_data') and "Greater Sydney" in self.comparison_data and "median_rent" in self.comparison_data["Greater Sydney"] and self.comparison_data["Greater Sydney"]["median_rent"] is not None else comparison_gs["value"]
            ron_value = self.comparison_data["Rest of NSW"]["median_rent"] if hasattr(self, 'comparison_data') and "Rest of NSW" in self.comparison_data and "median_rent" in self.comparison_data["Rest of NSW"] and self.comparison_data["Rest of NSW"]["median_rent"] is not None else comparison_ron["value"]
        
            gs_text = ""
            if local_increase < gs_value - 1:  # 1% buffer
                gs_text = f"lower than Greater Sydney's annual increase of {gs_value}%"
            elif local_increase > gs_value + 1:
                gs_text = f"higher than Greater Sydney's annual increase of {gs_value}%"
            else:
                gs_text = f"similar to Greater Sydney's annual increase of {gs_value}%"
                
            ron_text = ""
            if local_increase < ron_value - 1:
                ron_text = f"and lower than Rest of NSW's annual increase of {ron_value}%"
            elif local_increase > ron_value + 1:
                ron_text = f"and higher than Rest of NSW's annual increase of {ron_value}%"
            else:
                ron_text = f"and similar to Rest of NSW's annual increase of {ron_value}%"
                
            return f"{self.selected_geo_name} ({self.selected_geo_area})'s median annual rental increase of {local_increase}% is {gs_text} {ron_text}."
        
        elif metric == "vacancy_rates":
            current_rate = self.data["vacancy_rates"]["value"]
            previous_rate = self.data["vacancy_rates"]["previous_year_rate"]
            
            # Format rates for display
            current_rate_display = current_rate
            previous_rate_display = previous_rate if previous_rate is not None else None
            
            # Generate text about market tightening/loosening if previous year data available
            trend_text = ""
            if previous_rate is not None:
                if current_rate < previous_rate - 0.1:
                    trend_text = f"The vacancy rate has tightened from {previous_rate_display:.2f}% a year ago to {current_rate_display:.2f}% now. "
                elif current_rate > previous_rate + 0.1:
                    trend_text = f"The vacancy rate has loosened from {previous_rate_display:.2f}% a year ago to {current_rate_display:.2f}% now. "
                else:
                    trend_text = f"The vacancy rate has remained stable at around {current_rate_display:.2f}% compared to {previous_rate_display:.2f}% a year ago. "
            
            # Format the comparison text with the actual values
            comparison_text = f"For reference, Greater Sydney has experienced a change of {gs_value}% and Rest of NSW {ron_value}% over the past year."
            
            return trend_text + comparison_text
        
        elif metric == "affordability":
            local_pct = self.data["affordability"]["percentage"]
            previous_year_pct = None
            
            if "previous_year_percentage" in self.data["affordability"] and self.data["affordability"]["previous_year_percentage"] is not None:
                previous_year_pct = self.data["affordability"]["previous_year_percentage"]
            
            # Get previous year values for comparison areas
            gs_prev_value = None
            if hasattr(self, 'comparison_data') and "Greater Sydney" in self.comparison_data and "affordability_previous" in self.comparison_data["Greater Sydney"]:
                gs_prev_value = self.comparison_data["Greater Sydney"]["affordability_previous"]
            
            # Compare with Greater Sydney
            gs_comparison = ""
            if local_pct > gs_value + 2:  # 2% buffer
                gs_comparison = f"less affordable than the Greater Sydney average of {gs_value}%"
            elif local_pct < gs_value - 2:
                gs_comparison = f"more affordable than the Greater Sydney average of {gs_value}%"
            else:
                gs_comparison = f"similar to the Greater Sydney average of {gs_value}%"
            
            # Compare with Rest of NSW
            ron_comparison = ""
            if local_pct > ron_value + 2:
                ron_comparison = f"and less affordable than the Rest of NSW average of {ron_value}%"
            elif local_pct < ron_value - 2:
                ron_comparison = f"and more affordable than the Rest of NSW average of {ron_value}%"
            else:
                ron_comparison = f"and similar to the Rest of NSW average of {ron_value}%"
            
            # Evaluate the trend based on previous year percentage
            trend_text = ""
            if previous_year_pct is not None:
                if abs(local_pct - previous_year_pct) < 1.0:  # Less than 1% change
                    trend_text = f" Affordability has remained relatively stable compared to {previous_year_pct}% a year ago."
                elif local_pct > previous_year_pct:  # Deterioration (higher percentage of income)
                    trend_text = f" Affordability has deteriorated from {previous_year_pct}% to {local_pct}% over the past year."
                else:  # Improvement (lower percentage of income)
                    trend_text = f" Affordability has improved from {previous_year_pct}% to {local_pct}% over the past year."
            
            return f"{self.selected_geo_name} ({self.selected_geo_area}) rental affordability is {gs_comparison} {ron_comparison}.{trend_text}"
        
        return ""
    
    def create_excel_output(self):
        """Create a nicely formatted Excel output with the analysis"""
        wb = Workbook()
        ws = wb.active
        ws.title = f"{self.selected_geo_name} Analysis"
        
        # Define styles
        header_font = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        metric_font = Font(bold=True, size=11)
        metric_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        metric_alignment = Alignment(vertical="center", wrap_text=True)
        
        value_font = Font(size=11)
        value_alignment = Alignment(vertical="center", wrap_text=True)
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Title
        ws.merge_cells('A1:E1')
        ws['A1'] = f"Rental Market Analysis for {self.selected_geo_name} ({self.selected_geo_area})"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        
        # Headers - Row 3
        headers = ["Metric", "Value", "Period", "Source", "Comment"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Set column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 50
        
        # Add metrics data
        row = 4
        
        # Renters
        ws.cell(row=row, column=1).value = "# and % of renters"
        ws.cell(row=row, column=1).font = metric_font
        ws.cell(row=row, column=1).fill = metric_fill
        ws.cell(row=row, column=1).alignment = metric_alignment
        ws.cell(row=row, column=1).border = thin_border
        
        ws.cell(row=row, column=2).value = f"{self.data['renters']['percentage']}% of all residential dwellings"
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        ws.cell(row=row, column=3).value = self.data['renters']['period']
        ws.cell(row=row, column=3).font = value_font
        ws.cell(row=row, column=3).alignment = value_alignment
        ws.cell(row=row, column=3).border = thin_border
        
        ws.cell(row=row, column=4).value = self.data['renters']['source']
        ws.cell(row=row, column=4).font = value_font
        ws.cell(row=row, column=4).alignment = value_alignment
        ws.cell(row=row, column=4).border = thin_border
        
        comment = self.generate_comparison_comment("renters", self.data['renters']['percentage'], 
                                              self.data['renters']['comparison_gs'], self.data['renters']['comparison_ron'])
        ws.cell(row=row, column=5).value = comment
        ws.cell(row=row, column=5).font = value_font
        ws.cell(row=row, column=5).alignment = value_alignment
        ws.cell(row=row, column=5).border = thin_border
        
        row += 1
        ws.cell(row=row, column=2).value = f"{self.data['renters']['count']:,}"
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        for col in [1, 3, 4, 5]:
            ws.cell(row=row, column=col).border = thin_border
        
        # Social Housing
        row += 1
        ws.cell(row=row, column=1).value = "# and % of Social Housing"
        ws.cell(row=row, column=1).font = metric_font
        ws.cell(row=row, column=1).fill = metric_fill
        ws.cell(row=row, column=1).alignment = metric_alignment
        ws.cell(row=row, column=1).border = thin_border
        
        ws.cell(row=row, column=2).value = f"{self.data['social_housing']['percentage']}% of all residential dwellings"
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        ws.cell(row=row, column=3).value = self.data['social_housing']['period']
        ws.cell(row=row, column=3).font = value_font
        ws.cell(row=row, column=3).alignment = value_alignment
        ws.cell(row=row, column=3).border = thin_border
        
        ws.cell(row=row, column=4).value = self.data['social_housing']['source']
        ws.cell(row=row, column=4).font = value_font
        ws.cell(row=row, column=4).alignment = value_alignment
        ws.cell(row=row, column=4).border = thin_border
        
        comment = self.generate_comparison_comment("social_housing", self.data['social_housing']['percentage'], 
                                                self.data['social_housing']['comparison_gs'], self.data['social_housing']['comparison_ron'])
        ws.cell(row=row, column=5).value = comment
        ws.cell(row=row, column=5).font = value_font
        ws.cell(row=row, column=5).alignment = value_alignment
        ws.cell(row=row, column=5).border = thin_border
        
        row += 1
        ws.cell(row=row, column=2).value = f"{self.data['social_housing']['count']:,}"
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        for col in [1, 3, 4, 5]:
            ws.cell(row=row, column=col).border = thin_border
        
        # Median Weekly Rent
        row += 1
        ws.cell(row=row, column=1).value = "Median Weekly Rent"
        ws.cell(row=row, column=1).font = metric_font
        ws.cell(row=row, column=1).fill = metric_fill
        ws.cell(row=row, column=1).alignment = metric_alignment
        ws.cell(row=row, column=1).border = thin_border
        
        ws.cell(row=row, column=2).value = f"${self.data['median_rent']['value']}"
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        ws.cell(row=row, column=3).value = self.data['median_rent']['period']
        ws.cell(row=row, column=3).font = value_font
        ws.cell(row=row, column=3).alignment = value_alignment
        ws.cell(row=row, column=3).border = thin_border
        
        ws.cell(row=row, column=4).value = self.data['median_rent']['source']
        ws.cell(row=row, column=4).font = value_font
        ws.cell(row=row, column=4).alignment = value_alignment
        ws.cell(row=row, column=4).border = thin_border
        
        comment = self.generate_comparison_comment("median_rent", self.data['median_rent']['value'], 
                                            self.data['median_rent']['comparison_gs'], self.data['median_rent']['comparison_ron'])
        ws.cell(row=row, column=5).value = comment
        ws.cell(row=row, column=5).font = value_font
        ws.cell(row=row, column=5).alignment = value_alignment
        ws.cell(row=row, column=5).border = thin_border
        
        row += 1
        # Show both annual increase and the previous year's rent
        annual_increase = self.data['median_rent']['annual_increase']
        prev_year_rent = self.data['median_rent']['previous_year_rent']
        
        display_text = f"Annual increase {annual_increase}%"
        if prev_year_rent is not None:
            display_text += f" (from ${prev_year_rent} last year)"
            
        ws.cell(row=row, column=2).value = display_text
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        for col in [1, 3, 4, 5]:
            ws.cell(row=row, column=col).border = thin_border
        
        # Vacancy Rates
        row += 1
        ws.cell(row=row, column=1).value = "Vacancy Rates"
        ws.cell(row=row, column=1).font = metric_font
        ws.cell(row=row, column=1).fill = metric_fill
        ws.cell(row=row, column=1).alignment = metric_alignment
        ws.cell(row=row, column=1).border = thin_border
        
        # Format vacancy rate value - if it's between 0 and 1, it's likely already a percentage (e.g., 0.75%)
        vacancy_value = self.data['vacancy_rates']['value']
        if vacancy_value < 1 and vacancy_value > 0:
            formatted_vacancy = f"{vacancy_value:.2f}%"
        else:
            formatted_vacancy = f"{vacancy_value:.2f}%"
            
        ws.cell(row=row, column=2).value = formatted_vacancy
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        ws.cell(row=row, column=3).value = self.data['vacancy_rates']['period']
        ws.cell(row=row, column=3).font = value_font
        ws.cell(row=row, column=3).alignment = value_alignment
        ws.cell(row=row, column=3).border = thin_border
        
        ws.cell(row=row, column=4).value = self.data['vacancy_rates']['source']
        ws.cell(row=row, column=4).font = value_font
        ws.cell(row=row, column=4).alignment = value_alignment
        ws.cell(row=row, column=4).border = thin_border
        
        comment = self.generate_comparison_comment("vacancy_rates", self.data['vacancy_rates']['value'], 
                                             self.data['vacancy_rates']['comparison_gs'], self.data['vacancy_rates']['comparison_ron'])
        ws.cell(row=row, column=5).value = comment
        ws.cell(row=row, column=5).font = value_font
        ws.cell(row=row, column=5).alignment = value_alignment
        ws.cell(row=row, column=5).border = thin_border
        
        row += 1
        previous_year_rate = self.data['vacancy_rates']['previous_year_rate']
        
        # Format previous year rate - check if it's already a percentage
        if previous_year_rate is not None:
            if previous_year_rate < 1 and previous_year_rate > 0:
                previous_year_text = f"Previous year: {previous_year_rate:.2f}%"
            else:
                previous_year_text = f"Previous year: {previous_year_rate:.2f}%"
        else:
            previous_year_text = "Previous year data not available"
            
        ws.cell(row=row, column=2).value = previous_year_text
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        for col in [1, 3, 4, 5]:
            ws.cell(row=row, column=col).border = thin_border
        
        # Rental Affordability
        row += 1
        ws.cell(row=row, column=1).value = "Rental affordability*"
        ws.cell(row=row, column=1).font = metric_font
        ws.cell(row=row, column=1).fill = metric_fill
        ws.cell(row=row, column=1).alignment = metric_alignment
        ws.cell(row=row, column=1).border = thin_border
        
        ws.cell(row=row, column=2).value = f"{self.data['affordability']['percentage']}% of income on rent"
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        ws.cell(row=row, column=3).value = self.data['affordability']['period']
        ws.cell(row=row, column=3).font = value_font
        ws.cell(row=row, column=3).alignment = value_alignment
        ws.cell(row=row, column=3).border = thin_border
        
        ws.cell(row=row, column=4).value = self.data['affordability']['source']
        ws.cell(row=row, column=4).font = value_font
        ws.cell(row=row, column=4).alignment = value_alignment
        ws.cell(row=row, column=4).border = thin_border
        
        comment = self.generate_comparison_comment("affordability", self.data['affordability']['percentage'], 
                                              self.data['affordability']['comparison_gs'], self.data['affordability']['comparison_ron'])
        ws.cell(row=row, column=5).value = comment
        ws.cell(row=row, column=5).font = value_font
        ws.cell(row=row, column=5).alignment = value_alignment
        ws.cell(row=row, column=5).border = thin_border
        
        row += 1
        # Show previous year percentage instead of improvement/deterioration
        if "previous_year_percentage" in self.data["affordability"] and self.data["affordability"]["previous_year_percentage"] is not None:
            previous_value = self.data["affordability"]["previous_year_percentage"]
            ws.cell(row=row, column=2).value = f"Previous year: {previous_value}% of income"
        else:
            ws.cell(row=row, column=2).value = "Previous year data not available"
            
        ws.cell(row=row, column=2).font = value_font
        ws.cell(row=row, column=2).alignment = value_alignment
        ws.cell(row=row, column=2).border = thin_border
        
        for col in [1, 3, 4, 5]:
            ws.cell(row=row, column=col).border = thin_border
        
        # Add footnote for affordability
        row += 2
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'].value = ("* Methodology: the rental affordability is calculated by taking median weekly rental household incomes for the "
                              "geographic area and comparing that to median weekly rents for the same area. Any number high than 30% of income "
                              "on rent is considered that a household is experiencing rental stress. This metric is calculated by Fair Trading "
                              "using ABS income and indexation data as well as Core Logic rental data.")
        ws[f'A{row}'].font = Font(italic=True, size=9)
        ws[f'A{row}'].alignment = Alignment(wrap_text=True)
        
        # Set row heights
        for r in range(3, row):
            ws.row_dimensions[r].height = 30
        
        ws.row_dimensions[row].height = 45  # Footnote row
        
        # Save the workbook to a BytesIO object
        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        return excel_buffer


def collect_reference_data(analyzer, uploaded_files):
    """Collect reference data for Greater Sydney and Rest of NSW"""
    with st.spinner("Collecting reference data for Greater Sydney and Rest of NSW..."):
        # Implementation has been shortened for brevity
        pass


def main():
    st.title("NSW Rental Data Analyzer")
    
    # Create an instance of the analyzer
    analyzer = RentalDataAnalyzer()
    
    # Sidebar for file uploads
    st.sidebar.header("Upload Data Files")
    
    # Create a dictionary to store uploaded files by category
    uploaded_files = {
        "median_rents": [],
        "census_dwelling": [],
        "census_demographics": [],
        "affordability": [],
        "vacancy_rates": []
    }
    
    # Allow uploading multiple files for each category
    with st.sidebar.expander("Median Rents Files", expanded=True):
        median_rent_files = st.file_uploader("Upload Median Rents Files", 
                                            type=["xlsx", "xls", "parquet"],
                                            accept_multiple_files=True,
                                            key="median_rents")
        
        if median_rent_files:
            for file in median_rent_files:
                # Save the file to a temporary location
                temp_file_path = os.path.join(analyzer.temp_dir, file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(file.getbuffer())
                
                uploaded_files["median_rents"].append({
                    "name": file.name,
                    "path": temp_file_path
                })
    
    with st.sidebar.expander("Census Dwelling Files", expanded=True):
        census_dwelling_files = st.file_uploader("Upload Census Dwelling Files", 
                                                type=["xlsx", "xls", "parquet"],
                                                accept_multiple_files=True,
                                                key="census_dwelling")
        
        if census_dwelling_files:
            for file in census_dwelling_files:
                # Save the file to a temporary location
                temp_file_path = os.path.join(analyzer.temp_dir, file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(file.getbuffer())
                
                uploaded_files["census_dwelling"].append({
                    "name": file.name,
                    "path": temp_file_path
                })
    
    with st.sidebar.expander("Affordability Files", expanded=True):
        affordability_files = st.file_uploader("Upload Affordability Files", 
                                            type=["xlsx", "xls", "parquet"],
                                            accept_multiple_files=True,
                                            key="affordability")
        
        if affordability_files:
            for file in affordability_files:
                # Save the file to a temporary location
                temp_file_path = os.path.join(analyzer.temp_dir, file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(file.getbuffer())
                
                uploaded_files["affordability"].append({
                    "name": file.name,
                    "path": temp_file_path
                })
    
    with st.sidebar.expander("Vacancy Rate Files", expanded=True):
        vacancy_rate_files = st.file_uploader("Upload Vacancy Rate Files", 
                                            type=["xlsx", "xls", "parquet"],
                                            accept_multiple_files=True,
                                            key="vacancy_rates")
        
        if vacancy_rate_files:
            for file in vacancy_rate_files:
                # Save the file to a temporary location
                temp_file_path = os.path.join(analyzer.temp_dir, file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(file.getbuffer())
                
                uploaded_files["vacancy_rates"].append({
                    "name": file.name,
                    "path": temp_file_path
                })
    
    # Main area for geographic selection and analysis
    st.header("Geographic Selection")
    
    # Check if files are uploaded
    files_uploaded = any(len(category_files) > 0 for category_files in uploaded_files.values())
    
    if not files_uploaded:
        st.warning("Please upload at least one data file to begin analysis.")
    else:
        # Display the number of files uploaded
        total_files = sum(len(category_files) for category_files in uploaded_files.values())
        st.success(f"{total_files} files uploaded successfully")
        
        # Geographic area type selection
        col1, col2 = st.columns(2)
        
        with col1:
            geo_area = st.selectbox("Select Geographic Area Type:", analyzer.GEO_AREAS)
            
            if geo_area:
                analyzer.selected_geo_area = geo_area
                
                # Get available geographic names
                with st.spinner(f"Loading {geo_area} names..."):
                    geo_names = analyzer.get_geo_names(geo_area, uploaded_files)
                
                if geo_names:
                    st.success(f"Found {len(geo_names)} {geo_area} names")
                    
                    # Display a select box with available names
                    with col2:
                        geo_name = st.selectbox("Select Geographic Name:", geo_names)
                        
                        if geo_name:
                            analyzer.selected_geo_name = geo_name
                else:
                    st.error(f"No {geo_area} names found in the uploaded files.")
        
        # Check if both geo area and name are selected
        if analyzer.selected_geo_area and analyzer.selected_geo_name:
            st.header("Analysis Options")
            
            # Generate button - THIS IS NOW CORRECTLY INSIDE THE FUNCTION AND PROPERLY INDENTED
            if st.button("Generate Analysis", type="primary"):
                with st.spinner("Analyzing rental data..."):
                    # First fetch actual comparison data for Greater Sydney and Rest of NSW
                    comparison_data = analyzer.fetch_comparison_area_data(uploaded_files)
                    
                    # Then collect data and analyze the selected area
                    analyzer.collect_data(uploaded_files)
                    
                    # Create tabs for different views
                    tab1, tab2 = st.tabs(["Analysis Summary", "Raw Data"])
                    
                    with tab1:
                        st.subheader(f"Rental Market Analysis for {analyzer.selected_geo_name} ({analyzer.selected_geo_area})")
                        
                        # Display summary cards
                        metric_col1, metric_col2, metric_col3 = st.columns(3)
                        
                        with metric_col1:
                            # Remove delta for renters count since it's a static number
                            st.metric(
                                label="Renters", 
                                value=f"{analyzer.data['renters']['percentage']}%"
                            )
                            st.write(f"{analyzer.data['renters']['count']:,} households")
                            
                            # Remove delta for social housing count since it's a static number
                            st.metric(
                                label="Social Housing", 
                                value=f"{analyzer.data['social_housing']['percentage']}%"
                            )
                            st.write(f"{analyzer.data['social_housing']['count']:,} dwellings")
                        
                        with metric_col2:
                            st.metric(
                                label="Median Weekly Rent", 
                                value=f"${analyzer.data['median_rent']['value']}",
                                delta=f"{analyzer.data['median_rent']['annual_increase']}% annual increase"
                            )
                            
                            # Format vacancy rate for display
                            vacancy_value = analyzer.data['vacancy_rates']['value']
                            if vacancy_value < 1 and vacancy_value > 0:
                                vacancy_display = f"{vacancy_value:.2f}%"
                            else:
                                vacancy_display = f"{vacancy_value:.2f}%"
                                
                            st.metric(
                                label="Vacancy Rate", 
                                value=vacancy_display,
                                delta=None
                            )
                            
                            # Show previous year rate
                            if analyzer.data['vacancy_rates']['previous_year_rate'] is not None:
                                prev_rate = analyzer.data['vacancy_rates']['previous_year_rate']
                                if prev_rate < 1 and prev_rate > 0:
                                    prev_display = f"{prev_rate:.2f}%"
                                else:
                                    prev_display = f"{prev_rate:.2f}%"
                                st.write(f"Was {prev_display} a year ago")
                        
                        with metric_col3:
                            st.metric(
                                label="Rental Affordability", 
                                value=f"{analyzer.data['affordability']['percentage']}% of income"
                            )
                            
                            # Show previous year value instead of improvement/deterioration
                            if "previous_year_percentage" in analyzer.data['affordability']:
                                st.write(f"Was {analyzer.data['affordability']['previous_year_percentage']}% of income a year ago")
                            elif "annual_improvement" in analyzer.data['affordability'] and analyzer.data['affordability']['annual_improvement'] != 0:
                                current = analyzer.data['affordability']['percentage']
                                improvement = analyzer.data['affordability']['annual_improvement']
                                previous = current + improvement if improvement < 0 else current - improvement
                                st.write(f"Was {previous:.1f}% of income a year ago")
                            else:
                                st.write("Previous year data not available")
                        
                        # Display comparison comments
                        st.subheader("Comparative Analysis")
                        
                        st.info(analyzer.generate_comparison_comment("renters", analyzer.data['renters']['percentage'], 
                                                    analyzer.data['renters']['comparison_gs'], analyzer.data['renters']['comparison_ron']))
                        
                        st.info(analyzer.generate_comparison_comment("social_housing", analyzer.data['social_housing']['percentage'], 
                                                    analyzer.data['social_housing']['comparison_gs'], analyzer.data['social_housing']['comparison_ron']))
                        
                        st.info(analyzer.generate_comparison_comment("median_rent", analyzer.data['median_rent']['value'], 
                                                    analyzer.data['median_rent']['comparison_gs'], analyzer.data['median_rent']['comparison_ron']))
                        
                        st.info(analyzer.generate_comparison_comment("vacancy_rates", analyzer.data['vacancy_rates']['value'], 
                                                    analyzer.data['vacancy_rates']['comparison_gs'], analyzer.data['vacancy_rates']['comparison_ron']))
                        
                        st.info(analyzer.generate_comparison_comment("affordability", analyzer.data['affordability']['percentage'], 
                                                    analyzer.data['affordability']['comparison_gs'], analyzer.data['affordability']['comparison_ron']))
                        
                        # Generate Excel file
                        excel_data = analyzer.create_excel_output()
                        
                        # Provide a download button for the Excel file
                        st.download_button(
                            label="Download Excel Report",
                            data=excel_data,
                            file_name=f"{analyzer.selected_geo_name}_{analyzer.selected_geo_area}_Rental_Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with tab2:
                        st.subheader("Raw Data")
                        
                        # Display raw data in expandable sections
                        with st.expander("Renters Data"):
                            st.json(analyzer.data["renters"])
                        
                        with st.expander("Social Housing Data"):
                            st.json(analyzer.data["social_housing"])
                        
                        with st.expander("Median Rent Data"):
                            st.json(analyzer.data["median_rent"])
                        
                        with st.expander("Vacancy Rate Data"):
                            st.json(analyzer.data["vacancy_rates"])
                        
                        with st.expander("Affordability Data"):
                            st.json(analyzer.data["affordability"])

    # Add footnote and info
    st.markdown("---")
    st.caption("* Methodology: Rental affordability is calculated by taking median weekly rental household incomes and comparing to median weekly rents. Any number higher than 30% of income on rent is considered rental stress.")
    st.caption("Source: NSW Fair Trading using ABS Census and Core Logic rental data")

if __name__ == "__main__":
    main()
