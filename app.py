import streamlit as st
import pandas as pd
import io
import os
import glob
import hashlib
from datetime import datetime

st.set_page_config(
    page_title="Weekly Target Planner", 
    layout="wide",
    page_icon="📦",
    initial_sidebar_state="collapsed"
)

# =====================
# AUTHENTICATION SYSTEM
# =====================

def check_password():
    """Returns `True` if the user had the correct password."""
    
    def password_entered():
        """Checks whether the entered password is correct."""
        # Try to get password from Streamlit secrets (for cloud deployment)
        try:
            correct_password = st.secrets["app_password"]
        except:
            # Fallback to hardcoded password for local development
            correct_password = "Skycargo@123"
        
        entered_password = st.session_state["password"]
        
        if entered_password == correct_password:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password in session
        else:
            st.session_state["password_correct"] = False

    # Return True if password is validated
    if st.session_state.get("password_correct", False):
        return True

    # Show login form
    st.markdown("""
    <div style='text-align: center; padding: 40px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;'>
        <h1 style='color: white; margin: 0;'>🔐 Weekly Target Planner</h1>
        <p style='color: white; margin: 10px 0 0 0; font-size: 18px;'>Please Login to Continue</p>
        <p style='color: white; margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;'>Authorized Users Only</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Create login form
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### 🔐 Access Required")
        
        # Simple password form
        st.text_input("Enter Password", type="password", key="password", placeholder="Enter the access password")
        
        if st.button("🔓 Enter", key="login_button", use_container_width=True):
            password_entered()
        
        # Show error message if login failed
        if st.session_state.get("password_correct") == False:
            st.error("😞 Incorrect password. Please try again.")

    return False

# Check authentication before loading the app
if not check_password():
    st.stop()

# Add welcome header with logout button
# Header with logout button
col1, col2 = st.columns([4, 1])

with col1:
    st.markdown("""
    <div style='text-align: center; padding: 20px; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;'>
        <h1 style='color: white; margin: 0;'>📦 Weekly Target Planner</h1>
        <p style='color: white; margin: 10px 0 0 0; font-size: 18px;'>Target Planning Tool Version 8</p>
        <p style='color: white; margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;'>Hamad Alhammadi</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("<br>", unsafe_allow_html=True)  # Add spacing
    if st.button("🔐 Logout", key="logout_button", help="Logout and return to password screen"):
        # Clear all session state
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

def initialize_session_state():
    """Initialize all session state variables with default values"""
    defaults = {
        "selected_station": None,
        "targets_data": None,
        "weekly_average": None,
        "weekly_data": None,
        "recommendations": None,
        "show_recommendations": False
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def get_available_stations():
    """Scan folder for Database - *.xlsx files and extract station names"""
    try:
        # Get current directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Look for Database - *.xlsx files
        pattern = os.path.join(current_dir, "Database - *.xlsx")
        files = glob.glob(pattern)
        
        stations = []
        for file_path in files:
            # Extract station name from filename
            filename = os.path.basename(file_path)
            # Remove "Database - " and ".xlsx"
            station_name = filename.replace("Database - ", "").replace(".xlsx", "")
            stations.append(station_name)
        
        return sorted(stations)  # Sort alphabetically
    except Exception as e:
        st.error(f"Error scanning for station files: {str(e)}")
        return []

def load_station_data(station_name):
    """Load data for the selected station"""
    try:
        # Construct filename
        current_dir = os.path.dirname(os.path.abspath(__file__))
        filename = f"Database - {station_name}.xlsx"
        file_path = os.path.join(current_dir, filename)
        
        if not os.path.exists(file_path):
            return None, None, f"Station file not found: {filename}"
        
        # Load Excel file
        xls = pd.ExcelFile(file_path)
        
        if len(xls.sheet_names) < 2:
            return None, None, f"Excel file must contain at least 2 sheets (Export and Weekly Average)"
        
        # Load Sheet 1: Export (Targets data)
        targets_data = pd.read_excel(xls, sheet_name=0)
        
        # Load Sheet 2: Weekly Average
        weekly_avg_data = pd.read_excel(xls, sheet_name=1)
        
        # Clean column names
        targets_data.columns = targets_data.columns.str.strip()
        weekly_avg_data.columns = weekly_avg_data.columns.str.strip()
        
        # Validate required columns in targets
        required_targets_cols = ['Week', 'Tgt Wt', 'Trgt Yield', 'Tgt Rev']
        missing_cols = [col for col in required_targets_cols if col not in targets_data.columns]
        if missing_cols:
            return None, None, f"Missing required columns in Export sheet: {missing_cols}"
        
        # Validate required columns in weekly average
        required_weekly_cols = ['Week', 'Tonnage', 'Revenue', 'Yield']
        missing_cols = [col for col in required_weekly_cols if col not in weekly_avg_data.columns]
        if missing_cols:
            return None, None, f"Missing required columns in Weekly Average sheet: {missing_cols}"
        
        # Clean numeric data
        numeric_cols = ["Tonnage", "Revenue", "Yield"]
        for col in numeric_cols:
            if col in weekly_avg_data.columns:
                weekly_avg_data[col] = pd.to_numeric(weekly_avg_data[col], errors='coerce').fillna(0)
        
        return targets_data, weekly_avg_data, None
        
    except Exception as e:
        return None, None, f"Error loading station data: {str(e)}"

def get_currency_config(currency):
    """Get currency configuration including rates and symbols"""
    config = {
        "AED": {"rate": 1.0, "symbol": "AED"},
        "USD": {"rate": 1/3.67, "symbol": "$"},
        "BHD": {"rate": 0.102, "symbol": "BHD"}
    }
    return config.get(currency, config["USD"])

def validate_data_availability():
    """Check if required data is loaded with detailed error information"""
    required_data = ["targets_data", "weekly_average"]
    missing = []
    
    for key in required_data:
        if st.session_state.get(key) is None:
            missing.append(key)
    
    if missing:
        missing_readable = [key.replace('_', ' ').title() for key in missing]
        st.warning(f"📁 Missing data: {', '.join(missing_readable)}. Please select a station first.")
        st.info("💡 Make sure your station Excel file contains both Export and Weekly Average sheets.")
        st.stop()
    return True

def create_metric_box(value, label, background_color="#eeeeee", text_color="black"):
    """Create a styled metric display box"""
    return f"""
    <div style='background-color:{background_color}; padding:20px; border-radius:10px; text-align:center;'>
        <div style='color:{text_color}; font-weight:bold;'>{label}</div>
        <div style='font-size:24px; font-weight:bold;'>{value}</div>
    </div>
    """

def clean_and_validate_data(df):
    """Clean and validate data ensuring consistency and removing invalid values.
    This function ensures revenue = tonnage × yield for all rows."""
    try:
        if df is None or df.empty:
            st.warning("⚠️ No data provided to clean and validate.")
            return pd.DataFrame()
            
        df_clean = df.copy()
        
        # Ensure required columns exist
        required_cols = ['Tonnage', 'Yield', 'Revenue']
        missing_cols = [col for col in required_cols if col not in df_clean.columns]
        if missing_cols:
            st.error(f"❌ Missing required columns: {missing_cols}")
            return df
        
        for idx, row in df_clean.iterrows():
            try:
                tonnage = float(row['Tonnage']) if pd.notna(row['Tonnage']) else 0
                yield_val = float(row['Yield']) if pd.notna(row['Yield']) else 0
                revenue = float(row['Revenue']) if pd.notna(row['Revenue']) else 0
                
                # Remove negative values - no negative business metrics allowed
                tonnage = max(0, tonnage)
                yield_val = max(0, yield_val)
                revenue = max(0, revenue)
                
                # Apply "all or nothing" rule: if any key metric is 0, all become 0
                if tonnage == 0 or yield_val == 0:
                    df_clean.loc[idx, 'Tonnage'] = 0
                    df_clean.loc[idx, 'Yield'] = 0
                    df_clean.loc[idx, 'Revenue'] = 0
                else:
                    # Valid data: ensure revenue = tonnage × yield (this is the key adjustment logic)
                    df_clean.loc[idx, 'Tonnage'] = round(tonnage, 0)
                    df_clean.loc[idx, 'Yield'] = round(yield_val, 2)
                    # Always calculate revenue from tonnage × yield to maintain consistency
                    df_clean.loc[idx, 'Revenue'] = round(tonnage * yield_val, 2)
            except Exception as row_error:
                st.warning(f"⚠️ Error processing row {idx}: {str(row_error)}")
                # Set problematic row to zeros
                df_clean.loc[idx, 'Tonnage'] = 0
                df_clean.loc[idx, 'Yield'] = 0
                df_clean.loc[idx, 'Revenue'] = 0
        
        return df_clean
        
    except Exception as e:
        st.error(f"❌ Error in data cleaning: {str(e)}")
        return df if df is not None else pd.DataFrame()

def calculate_smart_recommendations(week_df, target_tonnage, target_revenue, target_avg_yield):
    """Generate smart recommendations using intelligent distribution to achieve exact targets"""
    try:
        if week_df is None or week_df.empty:
            st.error("❌ No data provided for recommendations.")
            return pd.DataFrame()
            
        if target_tonnage <= 0 or target_revenue <= 0 or target_avg_yield <= 0:
            st.warning("⚠️ Invalid target values provided (must be positive).")
            return week_df.copy()
            
        recommendations_df = week_df.copy()
        
        # Ensure required columns exist
        required_cols = ['Tonnage', 'Yield', 'Revenue']
        missing_cols = [col for col in required_cols if col not in recommendations_df.columns]
        if missing_cols:
            st.error(f"❌ Missing required columns in data: {missing_cols}")
            return week_df.copy()
        
        # Get current valid data (exclude zero rows)
        valid_data = recommendations_df[
            (recommendations_df['Tonnage'] > 0) & 
            (recommendations_df['Yield'] > 0) & 
            (recommendations_df['Revenue'] > 0)
        ].copy()
        
        if valid_data.empty:
            # No valid data - create equal distribution among all agents
            num_agents = len(recommendations_df)
            if num_agents > 0:
                recommendations_df['Tonnage'] = target_tonnage / num_agents
                recommendations_df['Yield'] = target_avg_yield
                recommendations_df['Revenue'] = target_revenue / num_agents
            return clean_and_validate_data(recommendations_df)
        
        # SMART DISTRIBUTION APPROACH
        # Calculate each agent's contribution weight based on their current performance
        total_current_tonnage = valid_data['Tonnage'].sum()
        total_current_revenue = valid_data['Revenue'].sum()
        
        if total_current_tonnage == 0 or total_current_revenue == 0:
            st.warning("⚠️ No valid performance data found for weight calculation.")
            return week_df.copy()
        
        # Create distribution weights based on current performance
        valid_data['tonnage_weight'] = valid_data['Tonnage'] / total_current_tonnage
        valid_data['revenue_weight'] = valid_data['Revenue'] / total_current_revenue
        valid_data['combined_weight'] = (valid_data['tonnage_weight'] + valid_data['revenue_weight']) / 2
        
        # Distribute targets based on performance weights
        for idx, row in recommendations_df.iterrows():
            try:
                if idx in valid_data.index:
                    # Agent has valid data - distribute based on weight
                    weight = valid_data.loc[idx, 'combined_weight']
                    
                    # Distribute tonnage based on weight
                    agent_target_tonnage = target_tonnage * weight
                    
                    # Distribute revenue based on weight  
                    agent_target_revenue = target_revenue * weight
                    
                    # Calculate yield to achieve both tonnage and revenue targets
                    if agent_target_tonnage > 0:
                        calculated_yield = agent_target_revenue / agent_target_tonnage
                    else:
                        calculated_yield = target_avg_yield
                    
                    # Assign calculated values
                    recommendations_df.at[idx, 'Tonnage'] = agent_target_tonnage
                    recommendations_df.at[idx, 'Revenue'] = agent_target_revenue
                    recommendations_df.at[idx, 'Yield'] = calculated_yield
                else:
                    # Agent has no valid data - set to zero
                    recommendations_df.at[idx, 'Tonnage'] = 0
                    recommendations_df.at[idx, 'Revenue'] = 0
                    recommendations_df.at[idx, 'Yield'] = 0
            except Exception as agent_error:
                st.warning(f"⚠️ Error processing agent at index {idx}: {str(agent_error)}")
                recommendations_df.at[idx, 'Tonnage'] = 0
                recommendations_df.at[idx, 'Revenue'] = 0
                recommendations_df.at[idx, 'Yield'] = 0
        
        # Final validation and cleanup
        return clean_and_validate_data(recommendations_df)
        
    except Exception as e:
        st.error(f"❌ Error in smart recommendations calculation: {str(e)}")
        st.error("💡 Using original data instead.")
        return week_df.copy() if week_df is not None else pd.DataFrame()

initialize_session_state()

# --- STATION SELECTION ---
st.markdown("### 🏢 Select Station")

# Get available stations
available_stations = get_available_stations()

if not available_stations:
    st.error("❌ No station files found!")
    st.info("💡 Please add station files in format: 'Database - [STATION].xlsx' (e.g., 'Database - BAH.xlsx')")
    st.stop()

# Station selection dropdown
selected_station = st.selectbox(
    "Choose Station to Plan For:",
    available_stations,
    index=0,
    help="Select the station you want to create targets for"
)

# Load data when station changes
if selected_station != st.session_state.get("selected_station"):
    st.session_state.selected_station = selected_station
    
    with st.spinner(f"Loading {selected_station} data..."):
        targets_data, weekly_avg_data, error = load_station_data(selected_station)
        
        if error:
            st.error(f"❌ {error}")
            st.stop()
        else:
            st.session_state.targets_data = targets_data
            st.session_state.weekly_average = weekly_avg_data
            st.success(f"✅ {selected_station} data loaded successfully!")
            st.rerun()

# Display current station info
if st.session_state.selected_station:
    st.info(f"📊 Currently planning for: **{st.session_state.selected_station}** station")

st.markdown("---")

# =====================
# WEEKLY PLANNER
# =====================
st.header("📊 Weekly Target Planner")

# Validate data availability with proper error handling
try:
    validate_data_availability()
except Exception as e:
    st.error(f"❌ Data validation error: {str(e)}")
    st.stop()

# Currency & Week selection with error handling
try:
    currency = st.selectbox("Select Currency", ["AED", "USD", "BHD"], index=0)
    currency_config = get_currency_config(currency)
    currency_symbol = currency_config["symbol"]
    rate = currency_config["rate"]

    # Safe week extraction
    if st.session_state.targets_data is None or "Week" not in st.session_state.targets_data.columns:
        st.error("❌ No valid targets data found.")
        st.stop()
        
    weeks = st.session_state.targets_data["Week"].dropna().unique()
    if len(weeks) == 0:
        st.error("❌ No weeks found in targets data.")
        st.stop()
    
    # Convert weeks to integers, skip non-numeric values (e.g., 'Total')
    weeks_int = []
    for week in weeks:
        try:
            if pd.notna(week) and str(week).replace('.', '').isdigit():
                weeks_int.append(int(float(week)))
        except (ValueError, TypeError):
            continue
    
    if not weeks_int:
        st.error("❌ No valid numeric weeks found in targets data.")
        st.stop()
        
    weeks_int.sort()  # Sort weeks in ascending order
    week_selected = st.selectbox("Select Week", weeks_int)

    # Get targets for week with validation
    tgt_df = st.session_state.targets_data[st.session_state.targets_data["Week"] == week_selected]
    if tgt_df.empty:
        st.warning("⚠️ No target data found for the selected week.")
        st.stop()
    
    tgt = tgt_df.iloc[0]
    
    # Safely extract target values with defaults
    orig_ton = float(tgt.get("Tgt Wt", 0)) if pd.notna(tgt.get("Tgt Wt")) else 0
    orig_yld = float(tgt.get("Trgt Yield", 0)) if pd.notna(tgt.get("Trgt Yield")) else 0
    orig_rev = float(tgt.get("Tgt Rev", 0)) if pd.notna(tgt.get("Tgt Rev")) else 0
    
    # Convert targets to selected currency (assuming source is AED)
    aed_rate = 1.0
    conv_tgt_yld = orig_yld / aed_rate * rate
    conv_tgt_rev = orig_rev / aed_rate * rate

    # Initialize or reset weekly_data with validation
    current_weekly_data = st.session_state.weekly_data or {}
    if current_weekly_data.get("week") != week_selected:
        st.session_state.weekly_data = {
            "week": week_selected,
            "current_tonnage": 0.0,
            "current_yield": 0.0,
            "current_revenue": 0.0
        }
    data = st.session_state.weekly_data
    
except Exception as e:
    st.error(f"❌ Error in week/currency selection: {str(e)}")
    st.stop()

# --- Weekly Targets ---
st.markdown("### 🎯 Weekly Targets")
try:
    target_cols = st.columns(3)
    target_values = [
        f"{orig_ton:,.0f} kg",
        f"{currency_symbol} {conv_tgt_yld:.2f} / kg",
        f"{currency_symbol} {conv_tgt_rev:,.0f}"
    ]
    target_labels = ["Tonnage", "Yield", "Revenue"]
    
    for col, label, value in zip(target_cols, target_labels, target_values):
        col.markdown(create_metric_box(value, label, "#bbdefb", "#0d47a1"), unsafe_allow_html=True)
except Exception as e:
    st.error(f"❌ Error displaying targets: {str(e)}")

# --- Gap to Target ---
st.markdown("### 📉 Gap to Target")
try:
    gap_cols = st.columns(3)
    
    # Safely get current values with defaults
    current_tonnage = float(data.get('current_tonnage', 0))
    current_yield = float(data.get('current_yield', 0))
    current_revenue = float(data.get('current_revenue', 0))
    
    gaps = [
        current_tonnage - orig_ton,
        current_yield - conv_tgt_yld,
        current_revenue - conv_tgt_rev
    ]
    targets = [orig_ton, conv_tgt_yld, conv_tgt_rev]
    current_values = [current_tonnage, current_yield, current_revenue]
    gap_labels = ["Tonnage Gap", "Yield Gap", "Revenue Gap"]
    
    for col, label, gap_value, target_value, current_value in zip(gap_cols, gap_labels, gaps, targets, current_values):
        try:
            # Calculate percentage achievement
            if target_value > 0:
                percentage = (current_value / target_value) * 100
            else:
                percentage = 0
            
            # Format display value with percentage and proper +/- signs
            if "Tonnage" in label:
                # Show gap with + or - sign, no decimals for tonnage
                gap_display = f"+{gap_value:,.0f}" if gap_value >= 0 else f"{gap_value:,.0f}"
                display_value = f"{gap_display} kg ({percentage:.1f}%)"
            elif "Yield" in label:
                # Show gap with + or - sign, keep decimals for yield
                gap_display = f"+{gap_value:,.2f}" if gap_value >= 0 else f"{gap_value:,.2f}"
                display_value = f"{currency_symbol} {gap_display} ({percentage:.1f}%)"
            else:
                # Show gap with + or - sign, no decimals for revenue
                gap_display = f"+{gap_value:,.0f}" if gap_value >= 0 else f"{gap_value:,.0f}"
                display_value = f"{currency_symbol} {gap_display} ({percentage:.1f}%)"
            
            # Choose background color based on percentage achievement
            if percentage >= 95.0:  # 95% or higher = green (excellent)
                bg_color = "#c8e6c9"
                text_color = "#2e7d32"
            elif percentage >= 85.0:  # 85-94% = yellow (good)
                bg_color = "#fff9c4"
                text_color = "#f57c00"
            else:  # Below 85% = red (needs attention)
                bg_color = "#ffcdd2"
                text_color = "#c62828"
            
            col.markdown(create_metric_box(display_value, label, bg_color, text_color), unsafe_allow_html=True)
        except Exception as gap_error:
            col.error(f"Error in {label}: {str(gap_error)}")
except Exception as e:
    st.error(f"❌ Error calculating gaps: {str(e)}")

# --- Current Performance ---
st.markdown("### 📦 Current Performance")
try:
    perf_cols = st.columns(3)
    
    performance_values = [
        f"{current_tonnage:,.0f} kg",
        f"{currency_symbol} {current_yield:.2f} / kg",
        f"{currency_symbol} {current_revenue:,.0f}"
    ]
    performance_labels = ["Tonnage", "Yield", "Revenue"]
    
    for col, label, value in zip(perf_cols, performance_labels, performance_values):
        col.markdown(create_metric_box(value, label), unsafe_allow_html=True)
except Exception as e:
    st.error(f"❌ Error displaying performance: {str(e)}")

st.markdown("<br>", unsafe_allow_html=True)

# --- Action Buttons ---
_, btn_col1, btn_col2, btn_col3, btn_col4, btn_col5, _ = st.columns([0.5, 1.2, 1.2, 1.2, 1.2, 1.2, 0.5])

with btn_col1:
    if st.button("Recommend", key="recommend", help="Generate smart recommendations to achieve weekly targets"):
        try:
            if st.session_state.weekly_average is None:
                st.error("❌ No weekly average data available. Please upload data first.")
            else:
                wa = st.session_state.weekly_average.copy()
                week_df = wa[wa['Week'] == week_selected].copy()
                
                if not week_df.empty:
                    # Define targets with validation
                    target_tonnage = max(0, orig_ton)
                    target_revenue = max(0, conv_tgt_rev)
                    target_avg_yield = max(0, conv_tgt_yld)
                    
                    if target_tonnage == 0 or target_revenue == 0 or target_avg_yield == 0:
                        st.warning("⚠️ Some target values are zero. Recommendations may not be optimal.")
                    
                    # Generate smart recommendations
                    recommendations_df = calculate_smart_recommendations(
                        week_df, target_tonnage, target_revenue, target_avg_yield
                    )
                    
                    # Store recommendations
                    st.session_state.recommendations = recommendations_df
                    st.session_state.show_recommendations = True
                    
                    st.success("✅ Smart recommendations generated!")
                    st.rerun()
                    
                else:
                    st.error("❌ No weekly average data found for the selected week.")
                    
        except Exception as e:
            st.error(f"❌ Error generating recommendations: {str(e)}")

with btn_col2:
    if st.button("Adjust", key="adjust", help="Clean and adjust data ensuring mathematical consistency"):
        try:
            # Check if we're working with recommendations or weekly average
            if st.session_state.get("show_recommendations", False) and st.session_state.get("recommendations") is not None:
                # Clean and adjust recommendations table
                recommendations_df = st.session_state.recommendations.copy()
                cleaned_df = clean_and_validate_data(recommendations_df)
                st.session_state.recommendations = cleaned_df
                st.success("✅ Recommendations cleaned and validated.")
            else:
                # Clean and adjust weekly average table
                if st.session_state.weekly_average is not None:
                    wa = st.session_state.weekly_average.copy()
                    week_df = wa[wa['Week'] == week_selected].copy()
                    other_weeks = wa[wa['Week'] != week_selected]
                    
                    if not week_df.empty:
                        cleaned_week_df = clean_and_validate_data(week_df)
                        st.session_state.weekly_average = pd.concat([other_weeks, cleaned_week_df], ignore_index=True)
                        st.success("✅ Weekly average cleaned and validated.")
                    else:
                        st.warning("⚠️ No data found for the selected week to adjust.")
                else:
                    st.error("❌ No weekly average data available to adjust.")
            
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error during data adjustment: {str(e)}")

with btn_col3:
    if st.button("Apply", key="apply", help="Apply the current table values to update current performance metrics"):
        try:
            # Check if we're applying recommendations or weekly average
            if st.session_state.get("show_recommendations", False) and st.session_state.get("recommendations") is not None:
                # Apply recommendations
                recommendations_df = st.session_state.recommendations
                
                total_tonnage = recommendations_df['Tonnage'].sum()
                total_revenue = recommendations_df['Revenue'].sum()
                # Calculate average yield from all agents in the table
                valid_yields = recommendations_df[recommendations_df['Yield'] > 0]['Yield']
                avg_yield = valid_yields.mean() if len(valid_yields) > 0 else 0
                
                # Update current performance
                data['current_tonnage'] = total_tonnage
                data['current_revenue'] = total_revenue
                data['current_yield'] = avg_yield
                
                st.success("✅ Recommendations applied to current performance.")
            else:
                # Apply weekly average
                if st.session_state.weekly_average is not None:
                    wa = st.session_state.weekly_average
                    week_df = wa[wa['Week'] == week_selected]
                    
                    if not week_df.empty:
                        total_tonnage = week_df['Tonnage'].sum()
                        total_revenue = week_df['Revenue'].sum()
                        # Calculate average yield from all agents in the table
                        valid_yields = week_df[week_df['Yield'] > 0]['Yield']
                        avg_yield = valid_yields.mean() if len(valid_yields) > 0 else 0
                        
                        # Update current performance
                        data['current_tonnage'] = total_tonnage
                        data['current_revenue'] = total_revenue
                        data['current_yield'] = avg_yield
                        
                        st.success("✅ Weekly average applied to current performance.")
                    else:
                        st.warning("⚠️ No weekly average data found for the selected week.")
                else:
                    st.error("❌ No weekly average data available to apply.")
            
            st.rerun()  # Refresh to show updated values
            
        except Exception as e:
            st.error(f"❌ Error applying data: {str(e)}")

with btn_col4:
    if st.button("Reset", key="reset", help="Reset all current performance values back to zero"):
        try:
            # Reset current performance to zero
            data['current_tonnage'] = 0.0
            data['current_revenue'] = 0.0
            data['current_yield'] = 0.0
            
            st.success("✅ Current performance values reset to zero.")
            st.rerun()  # Refresh to show updated values
            
        except Exception as e:
            st.error(f"❌ Error resetting values: {str(e)}")

with btn_col5:
    st.markdown("**📥 Export**")
    st.caption("Export shown below")

# --- Editable Weekly Average Table ---
st.markdown("### 📈 Weekly Average")

try:
    # Show recommendations table if available, otherwise show weekly average
    if st.session_state.get("show_recommendations", False) and st.session_state.get("recommendations") is not None:
        st.info("📊 Showing recommendations to meet targets. Click 'Apply' to use these values.")
        
        # Display editable recommendations table
        edited_recommendations = st.data_editor(
            st.session_state.recommendations, 
            key='recommendations_editor', 
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Tonnage": st.column_config.NumberColumn(
                    "Tonnage",
                    help="Recommended tonnage in kg",
                    format="%.0f",
                    min_value=0
                ),
                "Revenue": st.column_config.NumberColumn(
                    "Revenue", 
                    help=f"Recommended revenue amount ({currency})",
                    format="%.0f",
                    min_value=0
                ),
                "Yield": st.column_config.NumberColumn(
                    "Yield",
                    help=f"Recommended yield per kg ({currency})", 
                    format="%.2f",
                    min_value=0
                )
            }
        )
        
        # Update recommendations in session state with edited data
        st.session_state.recommendations = edited_recommendations
        
        # Export and Back buttons
        col1, col2, _ = st.columns([2, 2, 6])
        with col1:
            if st.button("📋 Back to Weekly Average", key="back_to_weekly"):
                st.session_state.show_recommendations = False
                st.rerun()
        with col2:
            # Export current recommendations table
            if st.button("📥 Export Current Table", key="export_recommendations", help="Export the current recommendations table as Excel"):
                try:
                    export_data = edited_recommendations.copy()
                    
                    if not export_data.empty:
                        # Create Excel file in memory
                        output = io.BytesIO()
                        
                        # Export to Excel with better formatting
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            export_data.to_excel(writer, sheet_name='Recommendations', index=False)
                            
                            # Get the workbook and worksheet
                            workbook = writer.book
                            worksheet = writer.sheets['Recommendations']
                            
                            # Auto-adjust column widths
                            for column in worksheet.columns:
                                max_length = 0
                                column_letter = column[0].column_letter
                                for cell in column:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                adjusted_width = min(max_length + 2, 50)
                                worksheet.column_dimensions[column_letter].width = adjusted_width
                        
                        output.seek(0)
                        
                        # Generate filename
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        station = st.session_state.selected_station or "Unknown"
                        filename = f"recommendations_{station}_week{week_selected}_{timestamp}.xlsx"
                        
                        st.download_button(
                            label=f"📥 Download Recommendations",
                            data=output.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_recommendations"
                        )
                        
                        st.success(f"✅ Recommendations table ready for download!")
                    else:
                        st.warning("⚠️ No recommendations data to export.")
                        
                except Exception as e:
                    st.error(f"❌ Export failed: {str(e)}")
                    st.error("💡 Please try again or contact support.")
            
    elif st.session_state.weekly_average is not None:
        wa = st.session_state.weekly_average.copy()
        week_data = wa[wa['Week'] == week_selected]
        
        if not week_data.empty:
            # Make the table editable
            edited_data = st.data_editor(
                week_data, 
                key='weekly_avg_editor', 
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Tonnage": st.column_config.NumberColumn(
                        "Tonnage",
                        help="Tonnage in kg",
                        format="%.0f",
                        min_value=0
                    ),
                    "Revenue": st.column_config.NumberColumn(
                        "Revenue", 
                        help="Revenue amount",
                        format="%.0f",
                        min_value=0
                    ),
                    "Yield": st.column_config.NumberColumn(
                        "Yield",
                        help="Yield per kg", 
                        format="%.2f",
                        min_value=0
                    )
                }
            )
            
            # Update session state with edited data
            other_weeks = wa[wa['Week'] != week_selected]
            st.session_state.weekly_average = pd.concat([other_weeks, edited_data], ignore_index=True)
            
            # Export button for weekly average
            st.markdown("---")
            col1, _, _ = st.columns([2, 4, 4])
            with col1:
                if st.button("📥 Export Current Table", key="export_weekly_avg", help="Export the current weekly average table as Excel"):
                    try:
                        export_data = edited_data.copy()
                        
                        if not export_data.empty:
                            # Create Excel file in memory
                            output = io.BytesIO()
                            
                            # Export to Excel with better formatting
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                export_data.to_excel(writer, sheet_name=f'Week_{week_selected}', index=False)
                                
                                # Get the workbook and worksheet
                                workbook = writer.book
                                worksheet = writer.sheets[f'Week_{week_selected}']
                                
                                # Auto-adjust column widths
                                for column in worksheet.columns:
                                    max_length = 0
                                    column_letter = column[0].column_letter
                                    for cell in column:
                                        try:
                                            if len(str(cell.value)) > max_length:
                                                max_length = len(str(cell.value))
                                        except:
                                            pass
                                    adjusted_width = min(max_length + 2, 50)
                                    worksheet.column_dimensions[column_letter].width = adjusted_width
                            
                            output.seek(0)
                            
                            # Generate filename
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            station = st.session_state.selected_station or "Unknown"
                            filename = f"weekly_avg_{station}_week{week_selected}_{timestamp}.xlsx"
                            
                            st.download_button(
                                label=f"📥 Download Weekly Average",
                                data=output.getvalue(),
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_weekly_avg"
                            )
                            
                            st.success(f"✅ Weekly average table ready for download!")
                        else:
                            st.warning("⚠️ No weekly average data to export.")
                            
                    except Exception as e:
                        st.error(f"❌ Export failed: {str(e)}")
                        st.error("💡 Please try again or contact support.")
        else:
            st.info("📋 No weekly average data found for the selected week.")
            st.info("💡 You can use the Recommend button to generate sample data.")
    else:
        st.warning("⚠️ No weekly average data available.")
        st.info("💡 Please select a station first.")
        
except Exception as e:
    st.error(f"❌ Error displaying table: {str(e)}")
    st.error("💡 Please refresh the page or check your data format.")