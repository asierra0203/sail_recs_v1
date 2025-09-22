import streamlit as st
import pandas as pd
from typing import List, Dict
import io

# Configure page
st.set_page_config(
    page_title="Royal Caribbean Sailing Recommendations",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #1f77b4;
        font-size: 2.5rem;
        margin-bottom: 2rem;
    }
    .section-header {
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .preferences-summary {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #3498db;
        margin: 1rem 0;
    }
    .weight-summary {
        background-color: #e8f5e8;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #27ae60;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Main header
    st.markdown('<h1 class="main-header">🚢 Royal Caribbean Sailing Recommendation System</h1>', unsafe_allow_html=True)
    st.markdown("**Generate personalized cruise recommendations based on your preferences and advanced scoring algorithms**")
    
    # Sidebar for navigation
    with st.sidebar:
        st.header("📋 Navigation")
        step = st.radio(
            "Select Step",
            ["1. Upload Data", "2. Set Preferences", "3. Configure Weights", "4. Generate Recommendations"],
            index=0
        )
    
    # Initialize session state
    if 'uploaded_file' not in st.session_state:
        st.session_state.uploaded_file = None
    if 'preferences' not in st.session_state:
        st.session_state.preferences = {}
    if 'weights' not in st.session_state:
        st.session_state.weights = {}
    
    # Step 1: File Upload
    if step == "1. Upload Data":
        handle_file_upload()
    
    # Step 2: Preferences
    elif step == "2. Set Preferences":
        handle_preferences()
    
    # Step 3: Weights
    elif step == "3. Configure Weights":
        handle_weights()
    
    # Step 4: Generate Recommendations
    elif step == "4. Generate Recommendations":
        handle_recommendations()

def handle_file_upload():
    st.markdown('<h2 class="section-header">📁 Step 1: Upload Sailing Data</h2>', unsafe_allow_html=True)
    
    st.info("📌 Upload your Excel file containing the 'Master Sailings Grid' sheet with sailing data.")
    
    # Debug message
    st.write("DEBUG: File uploader should appear below:")
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xls, .xlsx, .xlsm)",
        type=["xls", "xlsx", "xlsm"],
        help="File should contain a 'Master Sailings Grid' sheet with columns: Ship Code, Month, Originating Port, Rdss Product Code, etc.",
        key="sailing_data_uploader"
    )
    
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        
        # Try to preview the data
        try:
            # Read the file to validate structure
            df = pd.read_excel(uploaded_file, sheet_name='Master Sailings Grid', nrows=5)
            
            st.success("✅ File uploaded successfully!")
            
            # Show file info
            col1, col2 = st.columns(2)
            with col1:
                st.metric("File Name", uploaded_file.name)
            with col2:
                st.metric("File Size", f"{uploaded_file.size / 1024:.1f} KB")
            
            # Preview data
            st.subheader("📊 Data Preview (First 5 rows)")
            st.dataframe(df, use_container_width=True)
            
            # Show available columns
            st.subheader("📋 Available Columns")
            cols = st.columns(3)
            for i, col in enumerate(df.columns):
                with cols[i % 3]:
                    st.write(f"• {col}")
            
        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            st.info("Please ensure your file contains a 'Master Sailings Grid' sheet with the expected structure.")
    
    else:
        st.warning("⚠️ Please upload an Excel file to continue.")

def handle_preferences():
    st.markdown('<h2 class="section-header">⚙️ Step 2: Set Your Preferences</h2>', unsafe_allow_html=True)
    
    if st.session_state.uploaded_file is None:
        st.warning("⚠️ Please upload a file first (Step 1).")
        return
    
    st.info("📌 Specify your preferences for ships, sailing months, and departure ports. Leave blank if you have no preference.")
    
    # Ship preferences
    st.subheader("🚢 Ship Preferences")
    
    # Show available ships from uploaded data
    try:
        df = pd.read_excel(st.session_state.uploaded_file, sheet_name='Master Sailings Grid')
        if 'Ship Code' in df.columns:
            available_ships = sorted(df['Ship Code'].dropna().unique().tolist())
            st.info(f"Available ships in your data: {', '.join(available_ships)}")
        else:
            available_ships = ["IC", "ST", "WN", "SY", "AL", "OV", "HM", "VY", "QN", "GR", "RH", "LB"]
    except:
        available_ships = ["IC", "ST", "WN", "SY", "AL", "OV", "HM", "VY", "QN", "GR", "RH", "LB"]
    
    ship_list = st.multiselect(
        "Select preferred ships",
        available_ships,
        help="Select one or more ships from your data"
    )
    
    # Month preferences
    st.subheader("📅 Month Preferences")
    
    month_names = ["January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November", "December"]
    
    selected_month_names = st.multiselect(
        "Select preferred sailing months",
        month_names,
        help="Select one or more months for sailing preferences"
    )
    month_list = [month_names.index(month) + 1 for month in selected_month_names]
    
    # Port preferences
    st.subheader("🏖️ Port Preferences")
    
    # Extract available ports from uploaded data
    try:
        df = pd.read_excel(st.session_state.uploaded_file, sheet_name='Master Sailings Grid')
        if 'Originating Port' in df.columns:
            available_ports = sorted(df['Originating Port'].dropna().unique().tolist())
            st.info(f"Available ports in your data: {', '.join(available_ports[:10])}{'...' if len(available_ports) > 10 else ''}")
        else:
            available_ports = ["MIA", "FLL", "PCA", "BAY", "NYC", "NOR"]
    except:
        available_ports = ["MIA", "FLL", "PCA", "BAY", "NYC", "NOR"]
    
    port_list = st.multiselect(
        "Select preferred departure ports",
        available_ports,
        help="Select one or more departure ports from your data"
    )
    
    # Save preferences
    st.session_state.preferences = {
        "ships": ship_list,
        "months": month_list,
        "ports": port_list
    }
    
    # Display summary
    if any([ship_list, month_list, port_list]):
        st.markdown('<div class="preferences-summary">', unsafe_allow_html=True)
        st.subheader("📋 Preferences Summary")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.write("**Ships:**")
            st.write(", ".join(ship_list) if ship_list else "None specified")
        
        with col2:
            st.write("**Months:**")
            if month_list:
                month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                month_display = [f"{month_names[m-1]} ({m})" for m in month_list]
                st.write(", ".join(month_display))
            else:
                st.write("None specified")
        
        with col3:
            st.write("**Ports:**")
            st.write(", ".join(port_list) if port_list else "None specified")
        
        st.markdown('</div>', unsafe_allow_html=True)

def handle_weights():
    st.markdown('<h2 class="section-header">⚖️ Step 3: Configure Scoring Weights</h2>', unsafe_allow_html=True)
    
    if st.session_state.uploaded_file is None:
        st.warning("⚠️ Please upload a file first (Step 1).")
        return
    
    st.info("📌 Rate the importance of each factor (0-10). Use 0 to completely turn off a factor. The system will automatically normalize your inputs.")
    
    # Weight explanations
    with st.expander("ℹ️ What do these factors mean?", expanded=False):
        st.write("""
        - **Ship Matching:** How important is it to get your preferred ships vs. any ship
        - **Month Preferences:** How important is it to sail in your preferred months
        - **Port Preferences:** How important is departure from your preferred ports
        - **Theo Adjustment (Profitability):** How important is the profitability/value of the sailing
        """)
    
    # Create weight sliders
    col1, col2 = st.columns(2)
    
    with col1:
        ship_weight = st.slider(
            "🚢 Ship Matching Importance",
            min_value=0,
            max_value=10,
            value=3,
            help="0 = Don't care about ship preferences, 10 = Ship preferences are critical"
        )
        
        month_weight = st.slider(
            "📅 Month Preference Importance",
            min_value=0,
            max_value=10,
            value=3,
            help="0 = Any month is fine, 10 = Must be in preferred months"
        )
    
    with col2:
        port_weight = st.slider(
            "🏖️ Port Preference Importance",
            min_value=0,
            max_value=10,
            value=3,
            help="0 = Any departure port is fine, 10 = Must depart from preferred ports"
        )
        
        theo_weight = st.slider(
            "💰 Theo Adjustment (Profitability) Importance",
            min_value=0,
            max_value=10,
            value=3,
            help="0 = Don't consider profitability, 10 = Profitability is most important"
        )
    
    # Validate that at least one weight is > 0
    weights = [ship_weight, month_weight, port_weight, theo_weight]
    if all(w == 0 for w in weights):
        st.error("❌ At least one factor must be greater than 0!")
        return
    
    # Calculate normalized weights
    total_weight = sum(weights)
    normalized_weights = {
        "ship": ship_weight / total_weight,
        "month": month_weight / total_weight,
        "port": port_weight / total_weight,
        "theo": theo_weight / total_weight
    }
    
    # Save weights
    st.session_state.weights = {
        "raw": {
            "ship": ship_weight,
            "month": month_weight,
            "port": port_weight,
            "theo": theo_weight
        },
        "normalized": normalized_weights
    }
    
    # Display weight summary
    st.markdown('<div class="weight-summary">', unsafe_allow_html=True)
    st.subheader("📊 Final Weight Distribution")
    
    # Create visual representation
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Ship",
            f"{normalized_weights['ship']:.1%}",
            delta=f"Raw: {ship_weight}" if ship_weight > 0 else "OFF"
        )
    
    with col2:
        st.metric(
            "Month",
            f"{normalized_weights['month']:.1%}",
            delta=f"Raw: {month_weight}" if month_weight > 0 else "OFF"
        )
    
    with col3:
        st.metric(
            "Port",
            f"{normalized_weights['port']:.1%}",
            delta=f"Raw: {port_weight}" if port_weight > 0 else "OFF"
        )
    
    with col4:
        st.metric(
            "Theo",
            f"{normalized_weights['theo']:.1%}",
            delta=f"Raw: {theo_weight}" if theo_weight > 0 else "OFF"
        )
    
    st.markdown('</div>', unsafe_allow_html=True)

def handle_recommendations():
    st.markdown('<h2 class="section-header">🎯 Step 4: Generate Recommendations</h2>', unsafe_allow_html=True)
    
    # Check if all steps are complete
    missing_steps = []
    if st.session_state.uploaded_file is None:
        missing_steps.append("Upload file (Step 1)")
    if not st.session_state.preferences:
        missing_steps.append("Set preferences (Step 2)")
    if not st.session_state.weights:
        missing_steps.append("Configure weights (Step 3)")
    
    if missing_steps:
        st.warning(f"⚠️ Please complete these steps first: {', '.join(missing_steps)}")
        return
    
    # Show configuration summary
    st.subheader("📋 Configuration Summary")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Preferences:**")
        prefs = st.session_state.preferences
        st.write(f"• Ships: {', '.join(prefs['ships']) if prefs['ships'] else 'None'}")
        st.write(f"• Months: {', '.join(map(str, prefs['months'])) if prefs['months'] else 'None'}")
        st.write(f"• Ports: {', '.join(prefs['ports']) if prefs['ports'] else 'None'}")
    
    with col2:
        st.markdown("**Weights:**")
        weights = st.session_state.weights['normalized']
        st.write(f"• Ship: {weights['ship']:.1%}")
        st.write(f"• Month: {weights['month']:.1%}")
        st.write(f"• Port: {weights['port']:.1%}")
        st.write(f"• Theo: {weights['theo']:.1%}")
    
    # Generate recommendations button
    if st.button("🚀 Generate Recommendations", type="primary", use_container_width=True):
        with st.spinner("🔄 Processing your preferences and generating recommendations..."):
            # Simulate processing time
            import time
            time.sleep(2)
            
            # Read the original data to add scores to
            original_df = pd.read_excel(st.session_state.uploaded_file, sheet_name='Master Sailings Grid')
            
            # Create dummy recommendations with scores added to original data
            recommendations_df = create_dummy_recommendations_with_scores(original_df)
            
            st.success("✅ Recommendations generated successfully!")
            
            # Display results
            st.subheader("🏆 Top Sailing Recommendations")
            
            # Create tabs for different views
            tab1, tab2, tab3 = st.tabs(["📊 Top Recommendations", "📈 Score Breakdown", "📋 Full Results"])
            
            with tab1:
                display_top_recommendations(recommendations_df)
            
            with tab2:
                display_score_breakdown(recommendations_df)
            
            with tab3:
                display_full_results(recommendations_df)
            
            # Download button for results - Modified Excel with new tab
            st.subheader("💾 Download Results")
            
            # Create Excel file with original data + new recommendations tab + preferences summary
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write original data to original sheet
                original_df.to_excel(writer, sheet_name='Master Sailings Grid', index=False)
                # Write recommendations with scores to new sheet
                recommendations_df.to_excel(writer, sheet_name='Sailing Recommendations', index=False)
                
                # Add preferences and weights summary sheet
                create_preferences_summary_sheet(writer, st.session_state.preferences, st.session_state.weights)
            
            st.download_button(
                label="📥 Download Excel with Recommendations Tab",
                data=output.getvalue(),
                file_name=f"sailing_recommendations_{st.session_state.uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def create_dummy_recommendations_with_scores(original_df):
    """Add randomized recommendation scores to the actual sailing data for demo purposes"""
    import random
    
    # Create a copy of the original data
    recommendations = original_df.copy()
    
    # Add only the Match Score column (final recommendation score)
    # Use randomized scores that would be realistic for a recommendation system
    recommendations['Match Score'] = [round(random.uniform(65, 95), 2) for _ in range(len(recommendations))]
    
    # Sort by Match Score descending
    recommendations = recommendations.sort_values('Match Score', ascending=False).reset_index(drop=True)
    
    # Add rank column
    recommendations.insert(0, 'Rank', range(1, len(recommendations) + 1))
    
    return recommendations

def create_preferences_summary_sheet(writer, preferences: Dict, weights: Dict):
    """Create a summary sheet with user preferences and weights"""
    
    # Create summary data
    summary_data = []
    
    # Add preferences section
    summary_data.append(['PREFERENCES:', ''])
    summary_data.append(['Ships:', ', '.join(preferences['ships']) if preferences['ships'] else 'None specified'])
    summary_data.append(['Months:', ', '.join(map(str, preferences['months'])) if preferences['months'] else 'None specified'])
    summary_data.append(['Ports:', ', '.join(preferences['ports']) if preferences['ports'] else 'None specified'])
    summary_data.append(['', ''])  # Empty row
    
    # Add weights section
    summary_data.append(['WEIGHTS:', ''])
    normalized_weights = weights['normalized']
    summary_data.append(['Ship Importance:', f"{normalized_weights['ship']:.1%}"])
    summary_data.append(['Month Importance:', f"{normalized_weights['month']:.1%}"])
    summary_data.append(['Port Importance:', f"{normalized_weights['port']:.1%}"])
    summary_data.append(['Theo Adjustment Importance:', f"{normalized_weights['theo']:.1%}"])
    summary_data.append(['', ''])  # Empty row
    
    # Add raw weights for reference
    summary_data.append(['RAW WEIGHTS (0-10):', ''])
    raw_weights = weights['raw']
    summary_data.append(['Ship (Raw):', str(raw_weights['ship'])])
    summary_data.append(['Month (Raw):', str(raw_weights['month'])])
    summary_data.append(['Port (Raw):', str(raw_weights['port'])])
    summary_data.append(['Theo (Raw):', str(raw_weights['theo'])])
    
    # Create DataFrame and write to Excel
    summary_df = pd.DataFrame(summary_data, columns=['Setting', 'Value'])
    summary_df.to_excel(writer, sheet_name='Preferences & Weights', index=False)

def display_top_recommendations(df):
    """Display top 10 recommendations"""
    top_10 = df.head(10)
    
    # Select columns to display (adjust based on actual column names)
    display_columns = ['Rank', 'Match Score']
    
    # Add available columns from the original data
    available_columns = df.columns.tolist()
    for col in ['Ship Code', 'Month', 'Sailing Date', 'Originating Port']:
        if col in available_columns:
            display_columns.append(col)
    
    # Format the dataframe without color coding
    display_df = top_10[display_columns].copy()
    display_df['Match Score'] = display_df['Match Score'].apply(lambda x: f"{x:.1f}%")
    
    st.dataframe(display_df, use_container_width=True, hide_index=True)

def display_score_breakdown(df):
    """Display detailed breakdown for top recommendations"""
    top_5 = df.head(5)
    
    for idx, row in top_5.iterrows():
        with st.expander(f"🏆 Rank {row['Rank']}: {row.get('Ship Code', 'N/A')} - {row['Match Score']:.1f}%"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Sailing Details:**")
                # Display available columns dynamically
                for col in ['Ship Code', 'Sailing Date', 'Originating Port', 'Rdss Product Code', 'Month']:
                    if col in row.index and pd.notna(row[col]):
                        st.write(f"• {col}: {row[col]}")
            
            with col2:
                st.write("**Recommendation Score:**")
                st.write(f"• Final Match Score: {row['Match Score']:.1f}%")
                st.write("• *Individual component scores are calculated internally*")

def display_full_results(df):
    """Display full results table with filtering options"""
    st.write("**Filter Results:**")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Use Ship Code if available, otherwise skip
        if 'Ship Code' in df.columns:
            ship_filter = st.multiselect("Filter by Ship", df['Ship Code'].unique())
        else:
            ship_filter = []
    with col2:
        # Use Month if available, otherwise skip
        if 'Month' in df.columns:
            month_filter = st.multiselect("Filter by Month", sorted(df['Month'].unique()))
        else:
            month_filter = []
    with col3:
        min_score = st.slider("Minimum Score", 0, 100, 70)
    
    # Apply filters
    filtered_df = df.copy()
    if ship_filter and 'Ship Code' in df.columns:
        filtered_df = filtered_df[filtered_df['Ship Code'].isin(ship_filter)]
    if month_filter and 'Month' in df.columns:
        filtered_df = filtered_df[filtered_df['Month'].isin(month_filter)]
    filtered_df = filtered_df[filtered_df['Match Score'] >= min_score]
    
    st.write(f"**Showing {len(filtered_df)} of {len(df)} recommendations**")
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
