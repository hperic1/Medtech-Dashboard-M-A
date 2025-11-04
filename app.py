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
                st.success(f"✅ Found data file at: {path}")
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
        
        st.success(f"✅ Loaded {len(ma_df)} M&A deals and {len(inv_df)} investment deals")
        
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
        else:
            return f"${value:.1f}M"
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
            name='Deal Value ($M)',
            marker_color='#1f77b4',
            text=[format_currency(v) for v in quarterly_data['Total_Value']],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: %{text}<br><extra></extra>'
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
                title='Deal Value ($M)',
                side='left',
                showgrid=True
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False
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
    """Create JP Morgan chart for a specific category"""
    try:
        quarters = ['Q1', 'Q2', 'Q3', 'Q4']
        
        # Example data - replace with actual data
        data_map = {
            'M&A': [12000, 15000, 13500, 14200],
            'Venture': [8500, 9200, 8800, 9500],
            'IPO': [1200, 1500, 1100, 1300],
            'Licensing': [3500, 4000, 3800, 4200]
        }
        
        values = data_map.get(category, [0, 0, 0, 0])
        
        fig = go.Figure()
        
        # Add bars for each quarter
        fig.add_trace(go.Bar(
            x=quarters,
            y=values,
            marker_color=color,
            text=[format_currency(v) for v in values],
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Deal Value: %{text}<br><extra></extra>'
        ))
        
        # Update layout
        fig.update_layout(
            title=f'{category} Activity',
            xaxis=dict(title='Quarter'),
            yaxis=dict(title='Deal Value ($M)'),
            hovermode='x unified',
            showlegend=False,
            height=350,
            margin=dict(t=50, b=50, l=50, r=50)
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
    page = st.sidebar.radio("Go to", ["Deal Activity", "JP Morgan Summary", "Add New Deal"])
    
    if page == "Deal Activity":
        show_deal_activity(ma_df, inv_df)
    elif page == "JP Morgan Summary":
        show_jp_morgan_summary()
    elif page == "Add New Deal":
        show_add_deal(ma_df, inv_df)

def show_deal_activity(ma_df, inv_df):
    """Display deal activity dashboard"""
    st.header("Deal Activity Dashboard")
    
    # Create two columns for split view
    col1, col2 = st.columns(2)
    
    with col1:
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
            st.dataframe(filtered_ma, use_container_width=True, height=400)
        
        with tab2:
            # Top 3 deals
            top_deals = filtered_ma.copy()
            top_deals['Deal_Value_Numeric'] = top_deals['Deal Value'].apply(
                lambda x: float(str(x).replace('$', '').replace('B', '').replace('M', '').replace(',', '')) 
                if x != 'Undisclosed' else 0
            )
            top_deals = top_deals.nlargest(3, 'Deal_Value_Numeric')
            
            for idx, row in top_deals.iterrows():
                # Format the deal value properly
                value = row['Deal_Value_Numeric']
                if value >= 1000:
                    formatted_value = f"${value/1000:.1f}B"
                elif value > 0:
                    formatted_value = f"${value:.1f}M"
                else:
                    formatted_value = "Undisclosed"
                
                # Display with larger text
                st.markdown(f"### {row['Company']} ← {row['Acquirer']}")
                st.markdown(f"<h1 style='margin-top: -20px; color: #1f77b4;'>{formatted_value}</h1>", unsafe_allow_html=True)
                st.markdown(f"**{row['Deal Type (Merger / Acquisition)']}**")
                st.markdown("---")
        
        with tab3:
            fig = create_quarterly_chart(filtered_ma, 'Deal Value', 'M&A Activity by Quarter')
            if fig:
                st.plotly_chart(fig, use_container_width=True)
    
    with col2:
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
            st.dataframe(filtered_inv, use_container_width=True, height=400)
        
        with tab2:
            # Top 3 deals
            top_deals = filtered_inv.copy()
            top_deals['Amount_Numeric'] = top_deals['Amount Raised'].apply(
                lambda x: float(str(x).replace('$', '').replace('B', '').replace('M', '').replace(',', '')) 
                if x != 'Undisclosed' else 0
            )
            top_deals = top_deals.nlargest(3, 'Amount_Numeric')
            
            for idx, row in top_deals.iterrows():
                # Format the amount properly
                value = row['Amount_Numeric']
                if value >= 1000:
                    formatted_value = f"${value/1000:.1f}B"
                elif value > 0:
                    formatted_value = f"${value:.1f}M"
                else:
                    formatted_value = "Undisclosed"
                
                # Display with larger text
                st.markdown(f"### {row['Company']} - {row['Funding type (VC / PE)']}")
                st.markdown(f"<h1 style='margin-top: -20px; color: #ff7f0e;'>{formatted_value}</h1>", unsafe_allow_html=True)
                st.markdown(f"**{row['Lead Investors']}**")
                st.markdown("---")
        
        with tab3:
            fig = create_quarterly_chart(filtered_inv, 'Amount Raised', 'Venture Investment by Quarter')
            if fig:
                st.plotly_chart(fig, use_container_width=True)

def show_jp_morgan_summary():
    """Display JP Morgan summary"""
    st.header("JP Morgan MedTech Industry Report")
    
    st.markdown("### 2024 YTD Activity by Category")
    
    # Create 2x2 grid for charts
    row1_col1, row1_col2 = st.columns(2)
    row2_col1, row2_col2 = st.columns(2)
    
    # Top left: M&A
    with row1_col1:
        fig_ma = create_jp_morgan_chart_by_category('M&A', '#1f77b4')
        if fig_ma:
            st.plotly_chart(fig_ma, use_container_width=True)
    
    # Top right: Venture
    with row1_col2:
        fig_venture = create_jp_morgan_chart_by_category('Venture', '#ff7f0e')
        if fig_venture:
            st.plotly_chart(fig_venture, use_container_width=True)
    
    # Bottom left: IPO
    with row2_col1:
        fig_ipo = create_jp_morgan_chart_by_category('IPO', '#2ca02c')
        if fig_ipo:
            st.plotly_chart(fig_ipo, use_container_width=True)
    
    # Bottom right: Licensing
    with row2_col2:
        fig_licensing = create_jp_morgan_chart_by_category('Licensing', '#d62728')
        if fig_licensing:
            st.plotly_chart(fig_licensing, use_container_width=True)
    
    # Key trends below the charts
    st.markdown("---")
    st.subheader("Key Market Trends")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### M&A Activity")
        st.info("""
        - Strong Q2 performance with $15.0B in deal value
        - Continued focus on cardiovascular and minimally invasive technologies
        - Strategic consolidation driving large-scale acquisitions
        """)
        
        st.markdown("#### Venture Capital")
        st.info("""
        - Steady growth throughout 2024 with Q4 peak at $9.5B
        - AI-enabled diagnostics and digital health platforms attracting significant investment
        - Series B and C rounds dominating the funding landscape
        """)
        
    with col2:
        st.markdown("#### IPO Market")
        st.info("""
        - Gradual recovery with Q2 showing strongest performance at $1.5B
        - Selective high-quality offerings gaining traction
        - Investor appetite returning for profitable medtech companies
        """)
        
        st.markdown("#### Licensing Deals")
        st.info("""
        - Q2 peak at $4.0B reflecting strong partnership activity
        - Breakthrough therapy designations driving deal flow
        - Increasing focus on novel therapeutic platforms
        """)

def show_add_deal(ma_df, inv_df):
    """Manual deal entry form"""
    st.header("Add New Deal")
    
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
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Acquirer': [acquirer],
                        'Deal Type (Merger / Acquisition)': [deal_type_ma],
                        'Technology/Description': [technology],
                        'Deal Value': [deal_value if deal_value else 'Undisclosed'],
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
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Funding type (VC / PE)': [funding_type],
                        'Technology/Description': [technology],
                        'Amount Raised': [amount if amount else 'Undisclosed'],
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

if __name__ == "__main__":
    main()