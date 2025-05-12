def main():
    st.title("NSW Rental Data Analyzer")
    st.markdown("This tool analyzes NSW rental data and generates comprehensive reports and visualizations.")
    
    # Create an instance of the analyzer
    analyzer = RentalDataAnalyzer()
    
    # Create sidebar tabs for different data loading methods
    tab_folder, tab_upload = st.sidebar.tabs(["Load From Folder", "Upload Files"])
    
    # Folder input method
    with tab_folder:
        st.sidebar.header("Load Data from Folder")
        folder_path = st.sidebar.text_input("Enter folder path containing data files:", 
                                          placeholder="e.g., C:/Users/username/RentalData")
        
        if st.sidebar.button("Scan Folder"):
            if folder_path:
                success = analyzer.scan_folder(folder_path)
                if success:
                    st.sidebar.success("Data files found and loaded!")
                    st.session_state['data_loaded'] = True
            else:
                st.sidebar.error("Please enter a valid folder path.")
    
    # File upload method
    with tab_upload:
        st.sidebar.header("Upload Data Files")
        
        st.sidebar.write("Upload files for each data category:")
        
        # Create expanders for each data category
        with st.sidebar.expander("Median Rents Files", expanded=True):
            uploaded_median_rent = st.file_uploader("Upload Median Rents Files", 
                                                 type=["xlsx", "xls", "parquet"],
                                                 accept_multiple_files=True,
                                                 key="median_rent_upload")
            
        with st.sidebar.expander("Census Dwelling Files", expanded=True):
            uploaded_census = st.file_uploader("Upload Census Dwelling Files", 
                                            type=["xlsx", "xls", "parquet"],
                                            accept_multiple_files=True,
                                            key="census_upload")
            
        with st.sidebar.expander("Vacancy Rate Files", expanded=True):
            uploaded_vacancy = st.file_uploader("Upload Vacancy Rate Files", 
                                            type=["xlsx", "xls", "parquet"],
                                            accept_multiple_files=True,
                                            key="vacancy_upload")
            
        with st.sidebar.expander("Affordability Files", expanded=True):
            uploaded_affordability = st.file_uploader("Upload Affordability Files", 
                                                  type=["xlsx", "xls", "parquet"],
                                                  accept_multiple_files=True,
                                                  key="affordability_upload")
        
        # Combine all uploaded files
        all_uploads = (uploaded_median_rent or []) + (uploaded_census or []) + \
                     (uploaded_vacancy or []) + (uploaded_affordability or [])
        
        if all_uploads and st.sidebar.button("Process Uploaded Files"):
            success = analyzer.process_uploaded_files(all_uploads)
            if success:
                st.sidebar.success("Files processed successfully!")
                st.session_state['data_loaded'] = True
    
    # Main area - only show if data is loaded
    if 'data_loaded' in st.session_state and st.session_state['data_loaded']:
        # Get available geo areas
        available_geo_areas = analyzer.get_available_geo_areas()
        
        if available_geo_areas:
            # Create selection area
            st.header("Select Geographic Area to Analyze")
            
            col1, col2 = st.columns(2)
            
            with col1:
                selected_geo_area = st.selectbox("Geographic Area Type:", available_geo_areas)
                
                if selected_geo_area:
                    # Get names for the selected geo area
                    geo_names = analyzer.get_geo_names(selected_geo_area)
                    
                    if geo_names:
                        with col2:
                            selected_geo_name = st.selectbox("Select Name:", geo_names)
                            
                            if selected_geo_name:
                                # Set selected area and name in the analyzer
                                analyzer.selected_geo_area = selected_geo_area
                                analyzer.selected_geo_name = selected_geo_name
                    else:
                        st.warning(f"No geographic names found for {selected_geo_area}.")
            
            # Add button to analyze the data
            if analyzer.selected_geo_area and analyzer.selected_geo_name:
                if st.button("Analyze Data", type="primary"):
                    with st.spinner(f"Analyzing data for {analyzer.selected_geo_name}..."):
                        data = analyzer.collect_data_for_area(analyzer.selected_geo_area, analyzer.selected_geo_name)
                    
                    if data:
                        st.session_state['analysis_data'] = data
                        st.session_state['analysis_complete'] = True
    
    # Display analysis results if available
    if 'analysis_complete' in st.session_state and st.session_state['analysis_complete']:
        st.header(f"Rental Analysis for {analyzer.selected_geo_name} ({analyzer.selected_geo_area})")
        
        # Create tabs for different views
        tab_summary, tab_charts, tab_data, tab_export = st.tabs([
            "Summary", "Charts", "Raw Data", "Export"
        ])
        
        with tab_summary:
            # Create 3-column layout for key metrics
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Renters
                st.subheader("Rental Households")
                
                renters_data = analyzer.data.get("renters", {})
                
                st.metric(
                    label=f"Renters ({renters_data.get('period', 'N/A')})",
                    value=f"{renters_data.get('percentage', 'N/A')}%",
                    delta=f"{renters_data.get('percentage', 0) - renters_data.get('comparison_gs', {}).get('value', 0):.1f}% vs Greater Sydney"
                )
                st.write(f"Number of rental households: {renters_data.get('count', 'N/A'):,}")
                
                # Social Housing
                st.subheader("Social Housing")
                
                social_housing_data = analyzer.data.get("social_housing", {})
                
                st.metric(
                    label=f"Social Housing ({social_housing_data.get('period', 'N/A')})",
                    value=f"{social_housing_data.get('percentage', 'N/A')}%",
                    delta=f"{social_housing_data.get('percentage', 0) - social_housing_data.get('comparison_gs', {}).get('value', 0):.1f}% vs Greater Sydney"
                )
                st.write(f"Number of social housing dwellings: {social_housing_data.get('count', 'N/A'):,}")
            
            with col2:
                # Median Rent
                st.subheader("Median Weekly Rent")
                
                rent_data = analyzer.data.get("median_rent", {})
                
                st.metric(
                    label=f"Weekly Rent ({rent_data.get('period', 'N/A')})",
                    value=f"${rent_data.get('value', 'N/A')}",
                    delta=f"{rent_data.get('annual_increase', 'N/A')}% annual increase"
                )
                if rent_data.get('previous_year_rent'):
                    st.write(f"Previous year: ${rent_data.get('previous_year_rent', 'N/A')}")
                
                # Vacancy Rates
                st.subheader("Vacancy Rates")
                
                vacancy_data = analyzer.data.get("vacancy_rates", {})
                vacancy_value = vacancy_data.get('value', 0)
                vacancy_display = vacancy_value * 100 if vacancy_value < 1 else vacancy_value
                
                previous_year_rate = vacancy_data.get('previous_year_rate')
                if previous_year_rate is not None:
                    prev_display = previous_year_rate * 100 if previous_year_rate < 1 else previous_year_rate
                    delta = f"{(vacancy_display - prev_display):.2f}% since last year"
                else:
                    delta = None
                
                st.metric(
                    label=f"Vacancy Rate ({vacancy_data.get('period', 'N/A')})",
                    value=f"{vacancy_display:.2f}%",
                    delta=delta,
                    delta_color="normal"
                )
            
            with col3:
                # Affordability
                st.subheader("Rental Affordability")
                
                affordability_data = analyzer.data.get("affordability", {})
                
                # Get previous year percentage for delta
                current_pct = affordability_data.get('percentage', 0)
                prev_pct = affordability_data.get('previous_year_percentage')
                
                if prev_pct is not None:
                    delta = f"{current_pct - prev_pct:.1f}% since last year"
                else:
                    delta = None
                
                st.metric(
                    label=f"Rental Affordability ({affordability_data.get('period', 'N/A')})",
                    value=f"{current_pct}% of income",
                    delta=delta,
                    delta_color="inverse"  # Lower is better for affordability
                )
                
                st.info("Households spending more than 30% of income on rent are considered to be experiencing rental stress.")
            
            # Display comparison text
            st.subheader("Comparative Analysis")
            
            for metric, comment_generator in [
                ("renters", lambda: analyzer.generate_comparison_text(
                    "renters", 
                    analyzer.data['renters']['percentage'],
                    analyzer.data['renters']['comparison_gs'],
                    analyzer.data['renters']['comparison_ron']
                )),
                ("social_housing", lambda: analyzer.generate_comparison_text(
                    "social_housing", 
                    analyzer.data['social_housing']['percentage'],
                    analyzer.data['social_housing']['comparison_gs'],
                    analyzer.data['social_housing']['comparison_ron']
                )),
                ("median_rent", lambda: analyzer.generate_comparison_text(
                    "median_rent", 
                    analyzer.data['median_rent']['value'],
                    analyzer.data['median_rent']['comparison_gs'],
                    analyzer.data['median_rent']['comparison_ron']
                )),
                ("vacancy_rates", lambda: analyzer.generate_comparison_text(
                    "vacancy_rates", 
                    analyzer.data['vacancy_rates']['value'],
                    analyzer.data['vacancy_rates']['comparison_gs'],
                    analyzer.data['vacancy_rates']['comparison_ron']
                )),
                ("affordability", lambda: analyzer.generate_comparison_text(
                    "affordability", 
                    analyzer.data['affordability']['percentage'],
                    analyzer.data['affordability']['comparison_gs'],
                    analyzer.data['affordability']['comparison_ron']
                ))
            ]:
                st.info(comment_generator())
        
        with tab_charts:
            # Display time series charts if available
            has_time_series = any(
                analyzer.data.get(metric, {}).get('time_series') 
                for metric in ['median_rent', 'vacancy_rates', 'affordability']
            )
            
            if has_time_series:
                for metric, title, y_label in [
                    ("median_rent", "Median Weekly Rent", "Rent ($)"),
                    ("vacancy_rates", "Vacancy Rate", "Rate (%)"),
                    ("affordability", "Rental Affordability", "% of Income")
                ]:
                    time_series = analyzer.data.get(metric, {}).get("time_series")
                    
                    if time_series:
                        st.subheader(title)
                        
                        # Convert to DataFrame for Plotly
                        df = pd.DataFrame(time_series)
                        df['date'] = pd.to_datetime(df['date'])
                        
                        # For vacancy rates, convert to percentage if needed
                        if metric == "vacancy_rates":
                            if df['value'].max() < 1:
                                df['value'] = df['value'] * 100
                        
                        # Create chart
                        fig = px.line(
                            df, 
                            x='date', 
                            y='value',
                            title=f"{title} for {analyzer.selected_geo_name}",
                            labels={'value': y_label, 'date': 'Date'}
                        )
                        
                        # Add horizontal line at 30% for affordability
                        if metric == "affordability":
                            fig.add_hline(y=30, line_dash="dash", line_color="red", 
                                       annotation_text="Rental Stress Threshold (30%)")
                        
                        # Display chart
                        st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No time series data available for visualization.")
        
        with tab_data:
            # Display raw data in expandable sections
            for label, data_key in [
                ("Rental Households", "renters"),
                ("Social Housing", "social_housing"),
                ("Median Rent", "median_rent"),
                ("Vacancy Rates", "vacancy_rates"),
                ("Rental Affordability", "affordability")
            ]:
                with st.expander(f"{label} Data"):
                    st.json(analyzer.data.get(data_key, {}))
        
        with tab_export:
            st.subheader("Export Data")
            
            # Generate Excel report
            if st.button("Generate Excel Report"):
                with st.spinner("Creating Excel report..."):
                    excel_buffer = analyzer.create_excel_report()
                
                # Offer download button
                st.download_button(
                    label="Download Excel Report",
                    data=excel_buffer,
                    file_name=f"{analyzer.selected_geo_name}_{analyzer.selected_geo_area}_Rental_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
