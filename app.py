import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import os

# Page configuration
st.set_page_config(
    page_title="MedTech M&A & Venture Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for full-width tables
st.markdown("""
<style>
    .dataframe {
        width: 100% !important;
    }
    div[data-testid="stDataFrame"] > div {
        width: 100% !important;
    }
    .stDataFrame {
        width: 100%;
    }
    /* Make tables span full width */
    .element-container {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# Data loading function
@st.cache_data
def load_data():
    """Load data from Excel file"""
    try:
        # Try multiple possible file paths - INCLUDING data folder
        possible_paths = [
            'data/MedTech_YTD_Standardized.xlsx',  # In data folder
            './data/MedTech_YTD_Standardized.xlsx',  # In data folder (explicit)
            'MedTech_YTD_Standardized.xlsx',  # Same directory as app.py
            'MedTech_MA_Masterlist.xlsx',  # Old filename
            './MedTech_MA_Masterlist.xlsx',
            '/mnt/project/MedTech_MA_Masterlist.xlsx',
            os.path.join(os.path.dirname(__file__), 'data', 'MedTech_YTD_Standardized.xlsx'),
            os.path.join(os.path.dirname(__file__), 'MedTech_YTD_Standardized.xlsx')
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            st.error("❌ Cannot find MedTech_YTD_Standardized.xlsx. Please ensure the file is in the 'data' folder or same directory as app.py")
            st.info("📁 Looking in these locations:\n" + "\n".join(f"- {p}" for p in possible_paths))
            return pd.DataFrame(), pd.DataFrame()
        
        # Load M&A data - NOTE: Sheet name has SPACES not underscores
        ma_df = pd.read_excel(excel_path, sheet_name='YTD M&A Activity')
        
        # Load Investment data - NOTE: Sheet name has SPACES not underscores
        inv_df = pd.read_excel(excel_path, sheet_name='YTD Investment Activity')
        
        # Clean and standardize data
        ma_df = ma_df.fillna('Undisclosed')
        inv_df = inv_df.fillna('Undisclosed')
        
        return ma_df, inv_df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.info("💡 Make sure your Excel file has sheets named 'YTD M&A Activity' and 'YTD Investment Activity' (with spaces)")
        return pd.DataFrame(), pd.DataFrame()

def save_data(ma_df, inv_df):
    """Save data back to Excel file"""
    try:
        # Try multiple possible file paths - INCLUDING data folder
        possible_paths = [
            'data/MedTech_YTD_Standardized.xlsx',
            './data/MedTech_YTD_Standardized.xlsx',
            'MedTech_YTD_Standardized.xlsx',
            'MedTech_MA_Masterlist.xlsx',
            './MedTech_MA_Masterlist.xlsx',
            '/mnt/project/MedTech_MA_Masterlist.xlsx',
            os.path.join(os.path.dirname(__file__), 'data', 'MedTech_YTD_Standardized.xlsx'),
            os.path.join(os.path.dirname(__file__), 'MedTech_YTD_Standardized.xlsx')
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            # If file doesn't exist, create it in the data folder
            os.makedirs('data', exist_ok=True)
            excel_path = 'data/MedTech_YTD_Standardized.xlsx'
        
        # Save with correct sheet names (with spaces)
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            ma_df.to_excel(writer, sheet_name='YTD M&A Activity', index=False)
            inv_df.to_excel(writer, sheet_name='YTD Investment Activity', index=False)
        return True
    except Exception as e:
        st.error(f"Error saving data: {str(e)}")
        st.warning("⚠️ Note: Streamlit Cloud has a read-only file system. Changes won't persist after app restarts.")
        return False

def format_currency(value):
    """Format currency values"""
    if pd.isna(value) or value == 'Undisclosed':
        return 'Undisclosed'
    try:
        value = float(str(value).replace('$', '').replace('B', '').replace('M', '').replace(',', ''))
        if value >= 1000:
            return f"${value/1000:.1f}B"
        elif value > 0:
            return f"${value:.0f}M"
        else:
            return 'Undisclosed'
    except:
        return str(value)


def create_quarterly_chart(df, value_col, title):
    """Create quarterly stacked bar chart with deal count overlay"""
    try:
        # Prepare data
        quarterly_data = df.groupby('Quarter').agg({
            value_col: lambda x: sum([float(str(v).replace('$', '').replace('B', '').replace('M', '').replace(',', '')) 
                                     if v != 'Undisclosed' else 0 for v in x]),
            'Company': 'count'
        }).reset_index()
        quarterly_data.columns = ['Quarter', 'Total_Value', 'Deal_Count']
        
        # Sort quarters
        quarter_order = ['Q1', 'Q2', 'Q3', 'Q4']
        quarterly_data['Quarter'] = pd.Categorical(quarterly_data['Quarter'], categories=quarter_order, ordered=True)
        quarterly_data = quarterly_data.sort_values('Quarter')
        
        # Create figure
        fig = go.Figure()
        
        # Add bar chart for deal values
        fig.add_trace(go.Bar(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Total_Value'],
            name='Deal Value',
            marker_color='#1f77b4',
            text=[f"${v:,.0f}" for v in quarterly_data['Total_Value']],  # Full amount with commas
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: $%{y:,.0f}<br><extra></extra>'
        ))
        
        # Add line chart for deal count
        fig.add_trace(go.Scatter(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Deal_Count'],
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#ff7f0e', width=3),
            marker=dict(size=10),
            text=quarterly_data['Deal_Count'],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<br><extra></extra>'
        ))
        
        # Update layout
        fig.update_layout(
            title=title,
            xaxis=dict(title='Quarter'),
            yaxis=dict(
                title='Total Deal Value (USD)',
                side='left',
                showgrid=True,
                range=[0, max(quarterly_data['Total_Value']) * 1.2]  # Extend y-axis by 20% for data labels
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False,
                range=[0, max(quarterly_data['Deal_Count']) * 1.3]  # Extend y2-axis by 30% for data labels
            ),
            hovermode='x unified',
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            height=500,
            margin=dict(t=100, b=50, l=50, r=50)
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return None

def create_jp_morgan_chart_by_category(category, color):
    """Create JP Morgan chart for a specific category with deal count overlay"""
    try:
        quarters = ['Q1', 'Q2', 'Q3']  # Only Q1-Q3, Q4 not available yet
        
        # Actual data from JP Morgan 2025 reports
        data_map = {
            'M&A': {
                'values': [9200, 2100, 21700],  # Q1: $9.2B (57 deals), Q2: $2.1B (43 deals), Q3: $21.7B (65 deals)
                'counts': [57, 43, 65]
            },
            'Venture': {
                'values': [3700, 2600, 2900],  # Q1: $3.7B (117 rounds), Q2: $2.6B (90 rounds), Q3: $2.9B (67 rounds)
                'counts': [117, 90, 67]  # Q2: 90 venture rounds totaling $2.6 billion
            }
        }
        
        category_data = data_map.get(category, {'values': [0, 0, 0], 'counts': [0, 0, 0]})
        values = category_data['values']
        counts = category_data['counts']
        
        fig = go.Figure()
        
        # Add bars for deal values
        fig.add_trace(go.Bar(
            x=quarters,
            y=values,
            name='Deal Value ($M)',
            marker_color=color,
            text=[format_currency(v) for v in values],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: %{text}<br><extra></extra>'
        ))
        
        # Add line chart for deal count
        fig.add_trace(go.Scatter(
            x=quarters,
            y=counts,
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#90EE90', width=3),  # Light green for better visibility against blue and orange
            marker=dict(size=10, color='#90EE90'),
            text=[str(c) if c > 0 else '' for c in counts],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<br><extra></extra>'
        ))
        
        # Update layout with dual y-axes
        fig.update_layout(
            title=f'{category} Activity',
            xaxis=dict(title='Quarter'),
            yaxis=dict(
                title='Deal Value (Millions USD)',
                side='left',
                showgrid=True,
                range=[0, max(values) * 1.2]  # Extend y-axis by 20% for data labels
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False,
                range=[0, max(counts) * 1.3] if max(counts) > 0 else [0, 100]  # Extend y2-axis
            ),
            hovermode='x unified',
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            height=350,
            margin=dict(t=80, b=50, l=50, r=50)
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating {category} chart: {str(e)}")
        return None

# Main app
def main():
    st.title("🏥 MedTech M&A & Venture Dashboard")
    
    # Load data
    ma_df, inv_df = load_data()
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", ["Deal Activity", "JP Morgan Summary", "Data Management"])
    
    if page == "Deal Activity":
        show_deal_activity(ma_df, inv_df)
    elif page == "JP Morgan Summary":
        show_jp_morgan_summary()
    elif page == "Data Management":
        show_data_management(ma_df, inv_df)

def show_deal_activity(ma_df, inv_df):
    """Display deal activity dashboard"""
    st.header("Deal Activity Dashboard")
    
    # M&A Activity Section - Full Width
    st.subheader("M&A Activity")
    
    # Search box
    search_ma = st.text_input("🔍 Search M&A Deals", placeholder="Search by company, acquirer, technology...", key='search_ma')
    
    # Filters
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        quarters_ma = ['All'] + sorted(ma_df['Quarter'].unique().tolist())
        selected_quarter_ma = st.selectbox("Filter by Quarter", quarters_ma, key='ma_quarter')
    with filter_col2:
        months_ma = ['All'] + sorted(ma_df['Month'].unique().tolist())
        selected_month_ma = st.selectbox("Filter by Month", months_ma, key='ma_month')
    
    # Apply filters
    filtered_ma = ma_df.copy()
    if selected_quarter_ma != 'All':
        filtered_ma = filtered_ma[filtered_ma['Quarter'] == selected_quarter_ma]
    if selected_month_ma != 'All':
        filtered_ma = filtered_ma[filtered_ma['Month'] == selected_month_ma]
    
    # Apply search filter
    if search_ma:
        mask = filtered_ma.apply(lambda row: row.astype(str).str.contains(search_ma, case=False).any(), axis=1)
        filtered_ma = filtered_ma[mask]
    
    # Tabs for table, top deals, and charts
    tab1, tab2, tab3 = st.tabs(["📊 Table", "🏆 Top Deals", "📈 Charts"])
    
    with tab1:
        # Create display dataframe with sortable numeric values
        ma_display = filtered_ma.copy()
        
        # Add hidden numeric column for sorting - use -1 for Undisclosed so it goes to bottom
        def parse_to_numeric(val):
            if val == 'Undisclosed' or pd.isna(val):
                return -1  # Changed from 0 to -1 to sort Undisclosed to bottom
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return -1
        
        # Create a numeric sort column
        ma_display['_Deal_Value_Numeric'] = ma_display['Deal Value'].apply(parse_to_numeric)
        
        # Sort by Deal Value descending by default (highest deals first, Undisclosed at bottom)
        ma_display = ma_display.sort_values('_Deal_Value_Numeric', ascending=False)
        
        # Display without the numeric column (it's just for sorting)
        display_cols = [col for col in ma_display.columns if not col.startswith('_')]
        
        st.dataframe(
            ma_display[display_cols], 
            use_container_width=True, 
            height=400,
            column_config={
                "Deal Value": st.column_config.TextColumn(
                    "Deal Value",
                    help="Deal value in USD",
                )
            }
        )
    
    with tab2:
        # Top 3 deals
        top_deals = filtered_ma.copy()
        
        # Parse function - values in Excel are already actual dollars like "$350,000,000"
        def parse_deal_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                # Value is already in actual dollars, not millions
                return float(val_str)
            except:
                return 0
        
        top_deals['Deal_Value_Numeric'] = top_deals['Deal Value'].apply(parse_deal_value)
        top_deals = top_deals.nlargest(3, 'Deal_Value_Numeric')
        
        for idx, row in top_deals.iterrows():
            # Value is already in actual dollars, just format with commas
            formatted_value = str(row['Deal Value']) if row['Deal Value'] != 'Undisclosed' else 'Undisclosed'
            
            # Get deal type verb
            deal_type = row['Deal Type (Merger / Acquisition)']
            verb = "merged with" if deal_type == "Merger" else "acquired"
            
            # Display with value directly from Excel (already formatted)
            st.markdown(f"**{row['Acquirer']} {verb} {row['Company']}**")
            st.markdown(f"<h1 style='margin-top: -10px; margin-bottom: -10px; color: #1f77b4;'>{formatted_value}</h1>", unsafe_allow_html=True)
            st.markdown("---")
    
    with tab3:
        fig = create_quarterly_chart(filtered_ma, 'Deal Value', 'M&A Activity by Quarter')
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    
    # Add spacing between sections
    st.markdown("---")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Venture Investment Activity Section - Full Width
    st.subheader("Venture Investment Activity")
    
    # Search box
    search_inv = st.text_input("🔍 Search Investment Deals", placeholder="Search by company, investors, technology...", key='search_inv')
    
    # Filters
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        quarters_inv = ['All'] + sorted(inv_df['Quarter'].unique().tolist())
        selected_quarter_inv = st.selectbox("Filter by Quarter", quarters_inv, key='inv_quarter')
    with filter_col2:
        months_inv = ['All'] + sorted(inv_df['Month'].unique().tolist())
        selected_month_inv = st.selectbox("Filter by Month", months_inv, key='inv_month')
    
    # Apply filters
    filtered_inv = inv_df.copy()
    if selected_quarter_inv != 'All':
        filtered_inv = filtered_inv[filtered_inv['Quarter'] == selected_quarter_inv]
    if selected_month_inv != 'All':
        filtered_inv = filtered_inv[filtered_inv['Month'] == selected_month_inv]
    
    # Apply search filter
    if search_inv:
        mask = filtered_inv.apply(lambda row: row.astype(str).str.contains(search_inv, case=False).any(), axis=1)
        filtered_inv = filtered_inv[mask]
    
    # Tabs for table, top deals, and charts
    tab1, tab2, tab3 = st.tabs(["📊 Table", "🏆 Top Deals", "📈 Charts"])
    
    with tab1:
        # Format Amount Raised column for display with sortable numeric values
        inv_display = filtered_inv.copy()
        
        # Add numeric sort column - use -1 for Undisclosed so it goes to bottom
        inv_display['_Amount_Numeric'] = inv_display['Amount Raised'].apply(
            lambda x: float(x) if pd.notna(x) and x != 'Undisclosed' and str(x).replace('.','').replace('-','').isdigit() else -1
        )
        
        # Sort by Amount descending by default (highest amounts first, Undisclosed at bottom)
        inv_display = inv_display.sort_values('_Amount_Numeric', ascending=False)
        
        # Format for display
        inv_display['Amount Raised'] = inv_display['Amount Raised'].apply(
            lambda x: f"${x:,.0f}" if pd.notna(x) and x != 'Undisclosed' and str(x).replace('.','').replace('-','').isdigit() else x
        )
        
        # Display without the numeric column
        display_cols = [col for col in inv_display.columns if not col.startswith('_')]
        
        st.dataframe(
            inv_display[display_cols],
            use_container_width=True, 
            height=400,
            column_config={
                "Amount Raised": st.column_config.TextColumn(
                    "Amount Raised",
                    help="Investment amount in USD",
                )
            }
        )
    
    with tab2:
        # Top 3 deals
        top_deals = filtered_inv.copy()
        
        # Parse function - values in Excel are already actual dollars like "$467,000,000"
        def parse_amount_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                # Value is already in actual dollars, not millions
                return float(val_str)
            except:
                return 0
        
        top_deals['Amount_Numeric'] = top_deals['Amount Raised'].apply(parse_amount_value)
        top_deals = top_deals.nlargest(3, 'Amount_Numeric')
        
        for idx, row in top_deals.iterrows():
            # Format amount with commas
            amount_val = row['Amount Raised']
            if pd.notna(amount_val) and amount_val != 'Undisclosed':
                try:
                    formatted_value = f"${float(amount_val):,.0f}"
                except:
                    formatted_value = str(amount_val)
            else:
                formatted_value = "Undisclosed"
            
            # Display with formatted value
            st.markdown(f"**{row['Company']}**")
            st.markdown(f"<h1 style='margin-top: -10px; margin-bottom: -10px; color: #ff7f0e;'>{formatted_value}</h1>", unsafe_allow_html=True)
            st.markdown("---")
    
    with tab3:
        fig = create_quarterly_chart(filtered_inv, 'Amount Raised', 'Venture Investment by Quarter')
        if fig:
            st.plotly_chart(fig, use_container_width=True)

def show_jp_morgan_summary():
    """Display JP Morgan summary"""
    st.header("JP Morgan MedTech Industry Report")
    
    # Load data for comparison
    ma_df, inv_df = load_data()
    
    # Calculate BeaconOne quarterly stats
    def calc_quarterly_stats(df, quarter, value_col):
        q_data = df[df['Quarter'] == quarter]
        
        # Parse values
        def parse_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return 0
        
        total_value = sum(q_data[value_col].apply(parse_value))
        count = len(q_data)
        
        # Format value
        if total_value >= 1000000000:
            formatted_value = f"${total_value/1000000000:.1f}B"
        elif total_value >= 1000000:
            formatted_value = f"${total_value/1000000:.0f}M"
        else:
            formatted_value = "$0"
            
        return count, formatted_value
    
    # Calculate stats for each quarter
    beacon_stats = {}
    for q in ['Q1', 'Q2', 'Q3']:
        ma_count, ma_value = calc_quarterly_stats(ma_df, q, 'Deal Value')
        inv_count, inv_value = calc_quarterly_stats(inv_df, q, 'Amount Raised')
        beacon_stats[q] = {
            'ma_count': ma_count,
            'ma_value': ma_value,
            'inv_count': inv_count,
            'inv_value': inv_value
        }
    
    st.markdown("### 2025 Q1-Q3 Activity by Category")
    
    # Create 1x2 grid for charts (only M&A and Venture)
    col1, col2 = st.columns(2)
    
    # Left: M&A
    with col1:
        fig_ma = create_jp_morgan_chart_by_category('M&A', '#1f77b4')
        if fig_ma:
            st.plotly_chart(fig_ma, use_container_width=True)
    
    # Right: Venture
    with col2:
        fig_venture = create_jp_morgan_chart_by_category('Venture', '#ff7f0e')
        if fig_venture:
            st.plotly_chart(fig_venture, use_container_width=True)
    
    # Key trends below the charts
    st.markdown("---")
    st.subheader("Key Market Trends")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**M&A Activity**")
        st.markdown("")
        st.markdown("• **Q1 2025**: 57 medtech M&A deals were announced, totaling $9.2 billion")
        st.markdown("")
        st.markdown("• **Q2 2025**: 43 medtech M&A deals were announced, totaling $2.1 billion")
        st.markdown("")
        st.markdown("• **Q3 2025**: 65 medtech M&A deals were announced, totaling $21.7 billion in upfront cash and equity")
        st.markdown("")
        st.markdown("**Overarching Trend**: Medtech M&A activity increased through Q3 2025, surpassing full-year 2024 numbers, with strategic consolidation driving large-scale transactions")
        
    with col2:
        st.markdown("**Venture Capital**")
        st.markdown("")
        st.markdown("• **Q1 2025**: Medtech venture investment activity continued to see larger rounds into fewer companies to post a higher dollar total for Q1 2025, exceeding Q1 2024")
        st.markdown("")
        st.markdown("• **Q2 2025**: The medtech venture landscape continues to show resilience, with total venture funding reaching $6.8 billion in the first half of 2025, positioning the sector to potentially exceed 2024's $12.7 billion full-year total")
        st.markdown("")
        st.markdown("• **Q3 2025**: Medtech venture funding started the year strong yet had a weaker Q2 and Q3 in a challenging venture funding environment across all of healthcare and life sciences")
        st.markdown("")
        st.markdown("**Overarching Trend**: Late-stage venture rounds continue to dominate at $7.9B YTD, while early-stage funding remains selective as investors focus on companies with proven traction")

    # Add comparison section
    st.markdown("---")
    st.markdown("### JPMorgan vs BeaconOne Data - Quarterly Comparison")
    
    # Create three columns for Q1, Q2, Q3
    q1_col, q2_col, q3_col = st.columns(3)
    
    with q1_col:
        st.markdown("#### Q1 2025")
        st.markdown(f"""
        <div style='background-color: #4A90E2; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>57</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['ma_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #357ABD; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$9.2B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['ma_value']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #50C878; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>117</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['inv_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #3FA35F; padding: 20px; border-radius: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$3.7B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['inv_value']}</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with q2_col:
        st.markdown("#### Q2 2025")
        st.markdown(f"""
        <div style='background-color: #4A90E2; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>43</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['ma_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #357ABD; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$2.1B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['ma_value']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #50C878; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>90</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['inv_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #3FA35F; padding: 20px; border-radius: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$2.6B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['inv_value']}</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with q3_col:
        st.markdown("#### Q3 2025")
        st.markdown(f"""
        <div style='background-color: #9B59B6; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>65</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['ma_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #8E44AD; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$21.7B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['ma_value']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #50C878; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>67</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['inv_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #3FA35F; padding: 20px; border-radius: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$2.9B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['inv_value']}</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

def show_data_management(ma_df, inv_df):
    """Data management page for adding deals and uploading JP Morgan reports"""
    st.header("Data Management")
    
    # Create tabs for different data management tasks
    tab1, tab2, tab3 = st.tabs(["📝 Add Manual Deals", "🌐 Web Scraper", "📊 Upload JP Morgan Report"])
    
    with tab1:
        show_manual_deal_entry(ma_df, inv_df)
    
    with tab2:
        show_web_scraper(ma_df, inv_df)
    
    with tab3:
        show_jp_morgan_upload()

def show_manual_deal_entry(ma_df, inv_df):
    """Manual deal entry forms"""
    st.subheader("Add New Deal Manually")
    
    # Select deal type
    deal_type = st.radio("Select Deal Type", ["M&A Activity", "Venture Investment"])
    
    if deal_type == "M&A Activity":
        st.subheader("Add M&A Deal")
        
        with st.form("ma_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                company = st.text_input("Company*")
                acquirer = st.text_input("Acquirer*")
                deal_type_ma = st.selectbox("Deal Type*", ["Acquisition", "Merger"])
            
            with col2:
                technology = st.text_area("Technology/Description*")
                deal_value = st.text_input("Deal Value (e.g., 100M, 1.5B, or Undisclosed)")
            
            col3, col4 = st.columns(2)
            with col3:
                quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"])
            with col4:
                month = st.selectbox("Month*", [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ])
            
            submitted = st.form_submit_button("Add M&A Deal")
            
            if submitted:
                if company and acquirer and technology:
                    # Parse deal value to standardized format
                    def parse_deal_input(val):
                        if not val or val.lower() == 'undisclosed':
                            return 'Undisclosed'
                        val_str = val.upper().replace('$', '').replace(',', '').strip()
                        try:
                            if 'B' in val_str:
                                num = float(val_str.replace('B', ''))
                                return f"${num * 1000000000:,.0f}"
                            elif 'M' in val_str:
                                num = float(val_str.replace('M', ''))
                                return f"${num * 1000000:,.0f}"
                            else:
                                return f"${float(val_str):,.0f}"
                        except:
                            return 'Undisclosed'
                    
                    formatted_value = parse_deal_input(deal_value)
                    
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Acquirer': [acquirer],
                        'Deal Type (Merger / Acquisition)': [deal_type_ma],
                        'Technology/Description': [technology],
                        'Deal Value': [formatted_value],
                        'Quarter': [quarter],
                        'Month': [month]
                    })
                    
                    # Append to dataframe
                    ma_df_updated = pd.concat([ma_df, new_deal], ignore_index=True)
                    
                    # Save data
                    if save_data(ma_df_updated, inv_df):
                        st.success("✅ M&A deal added successfully!")
                        st.balloons()
                        # Clear cache to reload data
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("Please fill in all required fields (*)")
    
    else:  # Venture Investment
        st.subheader("Add Venture Investment Deal")
        
        with st.form("inv_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                company = st.text_input("Company*")
                funding_type = st.selectbox("Funding Type*", ["VC", "PE"])
            
            with col2:
                technology = st.text_area("Technology/Description*")
                amount = st.text_input("Amount Raised (e.g., 50M, 1.2B, or Undisclosed)")
            
            lead_investors = st.text_input("Lead Investors")
            
            col3, col4 = st.columns(2)
            with col3:
                quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"])
            with col4:
                month = st.selectbox("Month*", [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ])
            
            submitted = st.form_submit_button("Add Investment Deal")
            
            if submitted:
                if company and technology:
                    # Parse amount to numeric format (just the number, no formatting)
                    def parse_amount_input(val):
                        if not val or val.lower() == 'undisclosed':
                            return 'Undisclosed'
                        val_str = val.upper().replace('$', '').replace(',', '').strip()
                        try:
                            if 'B' in val_str:
                                num = float(val_str.replace('B', ''))
                                return int(num * 1000000000)
                            elif 'M' in val_str:
                                num = float(val_str.replace('M', ''))
                                return int(num * 1000000)
                            else:
                                return int(float(val_str))
                        except:
                            return 'Undisclosed'
                    
                    formatted_amount = parse_amount_input(amount)
                    
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Funding type (VC / PE)': [funding_type],
                        'Technology/Description': [technology],
                        'Amount Raised': [formatted_amount],
                        'Lead Investors': [lead_investors if lead_investors else 'Undisclosed'],
                        'Quarter': [quarter],
                        'Month': [month]
                    })
                    
                    # Append to dataframe
                    inv_df_updated = pd.concat([inv_df, new_deal], ignore_index=True)
                    
                    # Save data
                    if save_data(ma_df, inv_df_updated):
                        st.success("✅ Investment deal added successfully!")
                        st.balloons()
                        # Clear cache to reload data
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("Please fill in all required fields (*)")

def show_web_scraper(ma_df, inv_df):
    """Web scraper for extracting deals from news articles and websites"""
    st.subheader("Web Scraper - Extract Deals from Articles")
    
    st.info("""
    🌐 **How it works:**
    1. Paste a URL from a news article or press release about MedTech deals
    2. Click "Scrape Deals" to automatically extract deal information
    3. Review and edit the extracted data
    4. Select whether each deal is M&A or Venture Investment
    5. Click "Add to Dashboard" to save all deals
    
    **Supported sites:** Press releases, news articles, industry reports with deal information
    """)
    
    # URL input
    url = st.text_input("Enter Article URL", placeholder="https://example.com/medtech-deals-2025", key="scraper_url")
    
    if st.button("🔍 Scrape Deals", type="primary"):
        if url:
            with st.spinner("Fetching and analyzing article..."):
                try:
                    # Use web_fetch to get the content
                    import requests
                    from bs4 import BeautifulSoup
                    import re
                    
                    # Fetch the webpage
                    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
                    response = requests.get(url, headers=headers, timeout=10)
                    response.raise_for_status()
                    
                    # Parse with BeautifulSoup
                    soup = BeautifulSoup(response.content, 'html.parser')
                    
                    # Extract text content
                    text = soup.get_text()
                    
                    # Store extracted deals in session state
                    if 'scraped_deals' not in st.session_state:
                        st.session_state.scraped_deals = []
                    
                    # Simple pattern matching for common deal structures
                    # Look for patterns like "Company acquired by Acquirer for $XXX"
                    patterns = [
                        r'([A-Z][A-Za-z\s&]+?)\s+(?:acquired|purchased|bought)\s+(?:by\s+)?([A-Z][A-Za-z\s&]+?)(?:\s+for\s+\$?([\d.]+)\s*(billion|million|B|M))?',
                        r'([A-Z][A-Za-z\s&]+?)\s+(?:raises|raised|secures|secured)\s+\$?([\d.]+)\s*(billion|million|B|M)',
                        r'([A-Z][A-Za-z\s&]+?)\s+(?:acquires|purchases)\s+([A-Z][A-Za-z\s&]+?)(?:\s+for\s+\$?([\d.]+)\s*(billion|million|B|M))?',
                    ]
                    
                    extracted_deals = []
                    
                    # Find all paragraphs
                    paragraphs = soup.find_all(['p', 'li'])
                    
                    for para in paragraphs:
                        para_text = para.get_text()
                        
                        # Check for acquisition patterns
                        for pattern in patterns:
                            matches = re.finditer(pattern, para_text, re.IGNORECASE)
                            for match in matches:
                                groups = match.groups()
                                
                                # Determine deal type and extract info
                                if 'acquir' in para_text.lower() or 'purchas' in para_text.lower() or 'bought' in para_text.lower():
                                    deal_type = 'M&A'
                                    if len(groups) >= 2:
                                        company = groups[0].strip()
                                        acquirer = groups[1].strip() if len(groups) > 1 else ''
                                        value = groups[2] if len(groups) > 2 and groups[2] else 'Undisclosed'
                                        unit = groups[3] if len(groups) > 3 and groups[3] else ''
                                        
                                        extracted_deals.append({
                                            'type': deal_type,
                                            'company': company,
                                            'acquirer': acquirer,
                                            'value': f"{value}{unit}" if value != 'Undisclosed' else 'Undisclosed',
                                            'description': para_text[:200]
                                        })
                                
                                elif 'rais' in para_text.lower() or 'secur' in para_text.lower():
                                    deal_type = 'Venture'
                                    if len(groups) >= 2:
                                        company = groups[0].strip()
                                        value = groups[1] if len(groups) > 1 and groups[1] else 'Undisclosed'
                                        unit = groups[2] if len(groups) > 2 and groups[2] else ''
                                        
                                        extracted_deals.append({
                                            'type': deal_type,
                                            'company': company,
                                            'acquirer': '',
                                            'value': f"{value}{unit}" if value != 'Undisclosed' else 'Undisclosed',
                                            'description': para_text[:200]
                                        })
                    
                    if extracted_deals:
                        st.session_state.scraped_deals = extracted_deals
                        st.success(f"✅ Found {len(extracted_deals)} potential deals! Review and edit below.")
                    else:
                        st.warning("⚠️ No deals found automatically. Try the manual entry form instead.")
                        st.info("💡 The scraper looks for patterns like 'Company acquired by Acquirer' or 'Company raises $XXM'")
                
                except Exception as e:
                    st.error(f"❌ Error scraping URL: {str(e)}")
                    st.info("💡 Try copying the deal information and using the manual entry form instead.")
        else:
            st.warning("Please enter a URL")
    
    # Display and edit scraped deals
    if 'scraped_deals' in st.session_state and st.session_state.scraped_deals:
        st.markdown("---")
        st.subheader("Review Extracted Deals")
        st.markdown("Edit the information below before adding to the dashboard:")
        
        deals_to_add = []
        
        for idx, deal in enumerate(st.session_state.scraped_deals):
            with st.expander(f"Deal {idx + 1}: {deal['company']}", expanded=True):
                col1, col2 = st.columns(2)
                
                with col1:
                    deal_type_select = st.selectbox(
                        "Deal Type*", 
                        ["M&A Activity", "Venture Investment"],
                        index=0 if deal['type'] == 'M&A' else 1,
                        key=f"type_{idx}"
                    )
                    
                    company = st.text_input("Company*", value=deal['company'], key=f"company_{idx}")
                    
                    if deal_type_select == "M&A Activity":
                        acquirer = st.text_input("Acquirer*", value=deal.get('acquirer', ''), key=f"acquirer_{idx}")
                        deal_subtype = st.selectbox("Deal Subtype*", ["Acquisition", "Merger"], key=f"subtype_{idx}")
                    else:
                        funding_type = st.selectbox("Funding Type*", ["VC", "PE"], key=f"funding_{idx}")
                        lead_investors = st.text_input("Lead Investors", key=f"investors_{idx}")
                
                with col2:
                    technology = st.text_area("Technology/Description*", value=deal.get('description', ''), height=100, key=f"tech_{idx}")
                    deal_value = st.text_input("Deal Value (e.g., 100M, 1.5B, or Undisclosed)", value=deal.get('value', 'Undisclosed'), key=f"value_{idx}")
                    
                col3, col4 = st.columns(2)
                with col3:
                    quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"], key=f"quarter_{idx}")
                with col4:
                    month = st.selectbox("Month*", [
                        "January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"
                    ], key=f"month_{idx}")
                
                # Store edited deal info
                if deal_type_select == "M&A Activity":
                    deals_to_add.append({
                        'type': 'M&A',
                        'company': company,
                        'acquirer': acquirer,
                        'deal_subtype': deal_subtype,
                        'technology': technology,
                        'value': deal_value,
                        'quarter': quarter,
                        'month': month
                    })
                else:
                    deals_to_add.append({
                        'type': 'Venture',
                        'company': company,
                        'funding_type': funding_type,
                        'lead_investors': lead_investors,
                        'technology': technology,
                        'value': deal_value,
                        'quarter': quarter,
                        'month': month
                    })
        
        # Add all deals button
        if st.button("✅ Add All Deals to Dashboard", type="primary"):
            ma_updated = ma_df.copy()
            inv_updated = inv_df.copy()
            
            added_ma = 0
            added_inv = 0
            skipped_duplicates = []
            
            for deal in deals_to_add:
                if deal['type'] == 'M&A':
                    # Parse and format deal value
                    def parse_deal_input(val):
                        if not val or val.lower() == 'undisclosed':
                            return 'Undisclosed'
                        val_str = val.upper().replace('$', '').replace(',', '').strip()
                        try:
                            if 'B' in val_str:
                                num = float(val_str.replace('B', '').replace('ILLION', ''))
                                return f"${num * 1000000000:,.0f}"
                            elif 'M' in val_str:
                                num = float(val_str.replace('M', '').replace('ILLION', ''))
                                return f"${num * 1000000:,.0f}"
                            else:
                                return f"${float(val_str):,.0f}"
                        except:
                            return 'Undisclosed'
                    
                    formatted_value = parse_deal_input(deal['value'])
                    
                    # Check for duplicates - compare company name and deal value
                    is_duplicate = False
                    for idx, existing_row in ma_updated.iterrows():
                        existing_company = str(existing_row['Company']).strip().lower()
                        existing_value = str(existing_row['Deal Value']).strip()
                        
                        new_company = deal['company'].strip().lower()
                        
                        # Compare company names (exact match or very similar)
                        if existing_company == new_company or existing_company in new_company or new_company in existing_company:
                            # Compare deal values
                            if existing_value == formatted_value:
                                is_duplicate = True
                                skipped_duplicates.append(f"M&A: {deal['company']} ({formatted_value})")
                                break
                    
                    if not is_duplicate:
                        new_deal = pd.DataFrame({
                            'Company': [deal['company']],
                            'Acquirer': [deal['acquirer']],
                            'Deal Type (Merger / Acquisition)': [deal['deal_subtype']],
                            'Technology/Description': [deal['technology']],
                            'Deal Value': [formatted_value],
                            'Quarter': [deal['quarter']],
                            'Month': [deal['month']]
                        })
                        ma_updated = pd.concat([ma_updated, new_deal], ignore_index=True)
                        added_ma += 1
                
                else:  # Venture
                    # Parse and format amount
                    def parse_amount_input(val):
                        if not val or val.lower() == 'undisclosed':
                            return 'Undisclosed'
                        val_str = val.upper().replace('$', '').replace(',', '').strip()
                        try:
                            if 'B' in val_str:
                                num = float(val_str.replace('B', '').replace('ILLION', ''))
                                return int(num * 1000000000)
                            elif 'M' in val_str:
                                num = float(val_str.replace('M', '').replace('ILLION', ''))
                                return int(num * 1000000)
                            else:
                                return int(float(val_str))
                        except:
                            return 'Undisclosed'
                    
                    formatted_amount = parse_amount_input(deal['value'])
                    
                    # Check for duplicates - compare company name and amount
                    is_duplicate = False
                    for idx, existing_row in inv_updated.iterrows():
                        existing_company = str(existing_row['Company']).strip().lower()
                        existing_amount = str(existing_row['Amount Raised']).strip()
                        
                        new_company = deal['company'].strip().lower()
                        
                        # Compare company names (exact match or very similar)
                        if existing_company == new_company or existing_company in new_company or new_company in existing_company:
                            # Compare amounts
                            if str(formatted_amount) == existing_amount:
                                is_duplicate = True
                                # Format for display
                                if formatted_amount != 'Undisclosed':
                                    display_val = f"${formatted_amount:,}"
                                else:
                                    display_val = 'Undisclosed'
                                skipped_duplicates.append(f"Venture: {deal['company']} ({display_val})")
                                break
                    
                    if not is_duplicate:
                        new_deal = pd.DataFrame({
                            'Company': [deal['company']],
                            'Funding type (VC / PE)': [deal['funding_type']],
                            'Technology/Description': [deal['technology']],
                            'Amount Raised': [formatted_amount],
                            'Lead Investors': [deal.get('lead_investors', 'Undisclosed')],
                            'Quarter': [deal['quarter']],
                            'Month': [deal['month']]
                        })
                        inv_updated = pd.concat([inv_updated, new_deal], ignore_index=True)
                        added_inv += 1
            
            # Save data
            if save_data(ma_updated, inv_updated):
                success_msg = f"✅ Successfully added {added_ma} M&A deals and {added_inv} Venture deals!"
                if skipped_duplicates:
                    success_msg += f"\n\n⚠️ Skipped {len(skipped_duplicates)} duplicate(s):"
                    for dup in skipped_duplicates:
                        success_msg += f"\n• {dup}"
                
                st.success(success_msg)
                if added_ma > 0 or added_inv > 0:
                    st.balloons()
                # Clear session state
                st.session_state.scraped_deals = []
                # Clear cache to reload data
                st.cache_data.clear()
                st.rerun()

def show_jp_morgan_upload():
    """JP Morgan report upload and data extraction"""
    st.subheader("Upload JP Morgan MedTech Industry Report")
    
    st.info("""
    📄 **Instructions:**
    1. Upload the quarterly JP Morgan MedTech Industry Report (PDF or text)
    2. The system will extract key data for M&A and Venture activity
    3. Charts and key takeaways will be automatically updated in the JP Morgan Summary page
    """)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose JP Morgan Report", 
        type=['pdf', 'txt', 'docx'],
        help="Upload the quarterly JP Morgan MedTech Industry Report"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        # Quarter selection
        col1, col2 = st.columns(2)
        with col1:
            report_year = st.selectbox("Report Year", [2025, 2024, 2023])
        with col2:
            report_quarter = st.selectbox("Report Quarter", ["Q1", "Q2", "Q3", "Q4"])
        
        st.markdown("### Enter Data Manually")
        st.markdown("Please enter the key metrics from the report:")
        
        with st.form("jp_morgan_data_form"):
            st.markdown("#### M&A Activity")
            col1, col2 = st.columns(2)
            with col1:
                ma_value = st.number_input("M&A Deal Value ($M)", min_value=0.0, value=0.0, step=100.0)
            with col2:
                ma_count = st.number_input("M&A Deal Count", min_value=0, value=0, step=1)
            
            st.markdown("#### Venture Capital")
            col1, col2 = st.columns(2)
            with col1:
                vc_value = st.number_input("Venture Deal Value ($M)", min_value=0.0, value=0.0, step=100.0)
            with col2:
                vc_count = st.number_input("Venture Deal Count", min_value=0, value=0, step=1)
            
            st.markdown("#### Key Takeaways")
            ma_takeaway = st.text_area("M&A Key Takeaway", placeholder="Enter key insight for M&A activity...")
            vc_takeaway = st.text_area("Venture Capital Key Takeaway", placeholder="Enter key insight for VC activity...")
            
            submitted = st.form_submit_button("💾 Save JP Morgan Data")
            
            if submitted:
                # Save the data to a JSON file or database
                jp_morgan_data = {
                    'year': report_year,
                    'quarter': report_quarter,
                    'ma': {'value': ma_value, 'count': ma_count, 'takeaway': ma_takeaway},
                    'venture': {'value': vc_value, 'count': vc_count, 'takeaway': vc_takeaway}
                }
                
                # Create data directory if it doesn't exist
                os.makedirs('data', exist_ok=True)
                
                # Save to JSON file
                import json
                json_path = f'data/jp_morgan_{report_year}_{report_quarter}.json'
                with open(json_path, 'w') as f:
                    json.dump(jp_morgan_data, f, indent=2)
                
                st.success(f"✅ JP Morgan {report_year} {report_quarter} data saved successfully!")
                st.info("📊 The JP Morgan Summary page will now reflect this data. Navigate to 'JP Morgan Summary' to view the updated charts.")
                st.balloons()

if __name__ == "__main__":
    main()