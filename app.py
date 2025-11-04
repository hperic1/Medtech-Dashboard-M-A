import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import os
from utils.data_loader import load_data, save_data, check_duplicates, format_currency
from utils.charts import create_volume_chart, create_top_deals_card
from utils.scraper import scrape_article, parse_deal_from_text

# Page configuration
st.set_page_config(
    page_title="MedTech M&A & Investment Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .deal-value {
        font-size: 2rem;
        font-weight: bold;
        color: #1f77b4;
    }
    .deal-company {
        font-size: 1.2rem;
        font-weight: 600;
        margin-top: 0.5rem;
    }
    div[data-testid="stDataFrame"] {
        font-size: 14px;
    }
</style>
""", unsafe_allow_html=True)

# Load data
@st.cache_data
def get_data():
    return load_data()

# Initialize session state
if 'ma_data' not in st.session_state or 'inv_data' not in st.session_state:
    ma_data, inv_data = get_data()
    st.session_state.ma_data = ma_data
    st.session_state.inv_data = inv_data

# Main title
st.markdown('<h1 class="main-header">🏥 MedTech M&A & Investment Dashboard</h1>', unsafe_allow_html=True)

# Sidebar navigation
tab_selection = st.sidebar.radio(
    "Navigate to:",
    ["📊 Deal Activity Dashboard", "📈 JP Morgan Summary", "➕ Data Entry & Management"]
)

# ==================== TAB 1: DEAL ACTIVITY DASHBOARD ====================
if tab_selection == "📊 Deal Activity Dashboard":
    st.header("Deal Activity Dashboard")
    
    # Create two columns for split view
    col1, col2 = st.columns(2)
    
    # ========== M&A ACTIVITY ==========
    with col1:
        st.subheader("🤝 M&A Activity")
        
        # Sub-tabs for M&A
        ma_tab1, ma_tab2, ma_tab3 = st.tabs(["📋 Table", "🏆 Top 3 Deals", "📊 Charts"])
        
        with ma_tab1:
            # Filters
            ma_quarters = ['All'] + sorted(st.session_state.ma_data['Quarter'].unique().tolist())
            ma_months = ['All'] + sorted(st.session_state.ma_data['Month'].unique().tolist())
            
            filter_col1, filter_col2 = st.columns(2)
            with filter_col1:
                selected_ma_quarter = st.selectbox("Filter by Quarter", ma_quarters, key="ma_quarter")
            with filter_col2:
                selected_ma_month = st.selectbox("Filter by Month", ma_months, key="ma_month")
            
            # Apply filters
            filtered_ma = st.session_state.ma_data.copy()
            if selected_ma_quarter != 'All':
                filtered_ma = filtered_ma[filtered_ma['Quarter'] == selected_ma_quarter]
            if selected_ma_month != 'All':
                filtered_ma = filtered_ma[filtered_ma['Month'] == selected_ma_month]
            
            # Display table
            st.dataframe(
                filtered_ma,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
            st.caption(f"Total Deals: {len(filtered_ma)}")
        
        with ma_tab2:
            # Top 3 Deals
            st.markdown("### 🏆 Top 3 M&A Deals by Value")
            create_top_deals_card(st.session_state.ma_data, "M&A")
        
        with ma_tab3:
            # Charts
            st.markdown("### 📊 M&A Activity by Quarter")
            fig_ma = create_volume_chart(st.session_state.ma_data, "M&A")
            st.plotly_chart(fig_ma, use_container_width=True)
    
    # ========== INVESTMENT ACTIVITY ==========
    with col2:
        st.subheader("💰 Investment Activity")
        
        # Sub-tabs for Investment
        inv_tab1, inv_tab2, inv_tab3 = st.tabs(["📋 Table", "🏆 Top 3 Deals", "📊 Charts"])
        
        with inv_tab1:
            # Filters
            inv_quarters = ['All'] + sorted(st.session_state.inv_data['Quarter'].unique().tolist())
            inv_months = ['All'] + sorted(st.session_state.inv_data['Month'].unique().tolist())
            
            filter_col1, filter_col2 = st.columns(2)
            with filter_col1:
                selected_inv_quarter = st.selectbox("Filter by Quarter", inv_quarters, key="inv_quarter")
            with filter_col2:
                selected_inv_month = st.selectbox("Filter by Month", inv_months, key="inv_month")
            
            # Apply filters
            filtered_inv = st.session_state.inv_data.copy()
            if selected_inv_quarter != 'All':
                filtered_inv = filtered_inv[filtered_inv['Quarter'] == selected_inv_quarter]
            if selected_inv_month != 'All':
                filtered_inv = filtered_inv[filtered_inv['Month'] == selected_inv_month]
            
            # Display table
            st.dataframe(
                filtered_inv,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
            st.caption(f"Total Deals: {len(filtered_inv)}")
        
        with inv_tab2:
            # Top 3 Deals
            st.markdown("### 🏆 Top 3 Investment Deals by Value")
            create_top_deals_card(st.session_state.inv_data, "Investment")
        
        with inv_tab3:
            # Charts
            st.markdown("### 📊 Investment Activity by Quarter")
            fig_inv = create_volume_chart(st.session_state.inv_data, "Investment")
            st.plotly_chart(fig_inv, use_container_width=True)

# ==================== TAB 2: JP MORGAN SUMMARY ====================
elif tab_selection == "📈 JP Morgan Summary":
    st.header("JP Morgan MedTech Report Summary")
    
    st.info("📝 Upload JP Morgan reports to populate this section automatically")
    
    # Placeholder data structure for JP Morgan report
    if 'jpm_data' not in st.session_state:
        st.session_state.jpm_data = {
            'quarters': ['Q1 2025', 'Q2 2025', 'Q3 2025'],
            'ma_volume': [20, 25, 24],
            'ma_amount': [12.5, 15.2, 18.3],
            'venture_volume': [30, 28, 29],
            'venture_amount': [2.5, 3.1, 2.8],
            'ipo_volume': [2, 1, 3],
            'ipo_amount': [0.5, 0.3, 0.8],
            'licensing_volume': [5, 7, 6],
            'licensing_amount': [1.2, 1.5, 1.4],
            'key_deals': [
                "Stryker acquires Inari Medical for $4.9B",
                "Thermo Fisher acquires Solventum for $4.1B",
                "Neko Health raises $260M Series B",
                "Boston Scientific acquires Nalu Medical for $533M",
                "Waters/BD Biosciences merger valued at $17.5B"
            ]
        }
    
    # Create YTD graph
    st.subheader("Year-to-Date Deal Activity Summary")
    
    fig = go.Figure()
    
    # M&A bars
    fig.add_trace(go.Bar(
        name='M&A Amount ($B)',
        x=st.session_state.jpm_data['quarters'],
        y=st.session_state.jpm_data['ma_amount'],
        marker_color='#1f77b4',
        text=[f"${x}B" for x in st.session_state.jpm_data['ma_amount']],
        textposition='outside',
        yaxis='y'
    ))
    
    # Venture bars
    fig.add_trace(go.Bar(
        name='Venture Amount ($B)',
        x=st.session_state.jpm_data['quarters'],
        y=st.session_state.jpm_data['venture_amount'],
        marker_color='#2ca02c',
        text=[f"${x}B" for x in st.session_state.jpm_data['venture_amount']],
        textposition='outside',
        yaxis='y'
    ))
    
    # IPO bars
    fig.add_trace(go.Bar(
        name='IPO Amount ($B)',
        x=st.session_state.jpm_data['quarters'],
        y=st.session_state.jpm_data['ipo_amount'],
        marker_color='#ff7f0e',
        text=[f"${x}B" for x in st.session_state.jpm_data['ipo_amount']],
        textposition='outside',
        yaxis='y'
    ))
    
    # Licensing bars
    fig.add_trace(go.Bar(
        name='Licensing Amount ($B)',
        x=st.session_state.jpm_data['quarters'],
        y=st.session_state.jpm_data['licensing_amount'],
        marker_color='#d62728',
        text=[f"${x}B" for x in st.session_state.jpm_data['licensing_amount']],
        textposition='outside',
        yaxis='y'
    ))
    
    # Volume line overlay
    total_volume = [
        st.session_state.jpm_data['ma_volume'][i] + 
        st.session_state.jpm_data['venture_volume'][i] + 
        st.session_state.jpm_data['ipo_volume'][i] + 
        st.session_state.jpm_data['licensing_volume'][i]
        for i in range(len(st.session_state.jpm_data['quarters']))
    ]
    
    fig.add_trace(go.Scatter(
        name='Total Deal Volume',
        x=st.session_state.jpm_data['quarters'],
        y=total_volume,
        mode='lines+markers+text',
        line=dict(color='#9467bd', width=3),
        marker=dict(size=10),
        text=[f"{x} deals" for x in total_volume],
        textposition='top center',
        yaxis='y2'
    ))
    
    fig.update_layout(
        barmode='group',
        height=600,
        xaxis=dict(title='Quarter'),
        yaxis=dict(title='Deal Amount ($B)', side='left'),
        yaxis2=dict(title='Deal Volume', overlaying='y', side='right'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        template='plotly_white'
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Key Deals
    st.subheader("🔑 Key Deals Highlighted in Report")
    for deal in st.session_state.jpm_data['key_deals']:
        st.markdown(f"• {deal}")

# ==================== TAB 3: DATA ENTRY & MANAGEMENT ====================
elif tab_selection == "➕ Data Entry & Management":
    st.header("Data Entry & Management")
    
    entry_method = st.radio(
        "Choose data entry method:",
        ["📰 Article Scraping", "✍️ Manual Entry", "📄 Upload JP Morgan Report"]
    )
    
    # ========== ARTICLE SCRAPING ==========
    if entry_method == "📰 Article Scraping":
        st.subheader("Scrape Article for Deals")
        
        article_input_method = st.radio(
            "Input method:",
            ["Paste URL", "Paste Text"]
        )
        
        if article_input_method == "Paste URL":
            url = st.text_input("Enter article URL:")
            if st.button("Scrape Article"):
                if url:
                    with st.spinner("Scraping article..."):
                        try:
                            deals = scrape_article(url)
                            if deals:
                                st.success(f"Found {len(deals)} potential deal(s)!")
                                
                                for i, deal in enumerate(deals):
                                    st.markdown(f"**Deal {i+1}:**")
                                    st.json(deal)
                                    
                                    if st.button(f"Add Deal {i+1}", key=f"add_{i}"):
                                        # Check for duplicates and add
                                        deal_type = st.selectbox(
                                            "Deal Type:",
                                            ["M&A", "Investment"],
                                            key=f"type_{i}"
                                        )
                                        
                                        if deal_type == "M&A":
                                            if not check_duplicates(deal, st.session_state.ma_data, "M&A"):
                                                new_row = pd.DataFrame([deal])
                                                st.session_state.ma_data = pd.concat([st.session_state.ma_data, new_row], ignore_index=True)
                                                save_data(st.session_state.ma_data, st.session_state.inv_data)
                                                st.success("Deal added to M&A data!")
                                                st.rerun()
                                            else:
                                                st.warning("Duplicate deal detected!")
                                        else:
                                            if not check_duplicates(deal, st.session_state.inv_data, "Investment"):
                                                new_row = pd.DataFrame([deal])
                                                st.session_state.inv_data = pd.concat([st.session_state.inv_data, new_row], ignore_index=True)
                                                save_data(st.session_state.ma_data, st.session_state.inv_data)
                                                st.success("Deal added to Investment data!")
                                                st.rerun()
                                            else:
                                                st.warning("Duplicate deal detected!")
                            else:
                                st.warning("No deals found in article")
                        except Exception as e:
                            st.error(f"Error scraping article: {str(e)}")
                else:
                    st.warning("Please enter a URL")
        
        else:  # Paste Text
            article_text = st.text_area("Paste article text:", height=300)
            if st.button("Extract Deals"):
                if article_text:
                    with st.spinner("Analyzing text..."):
                        try:
                            deals = parse_deal_from_text(article_text)
                            if deals:
                                st.success(f"Found {len(deals)} potential deal(s)!")
                                st.json(deals)
                            else:
                                st.warning("No deals found in text")
                        except Exception as e:
                            st.error(f"Error parsing text: {str(e)}")
                else:
                    st.warning("Please paste article text")
    
    # ========== MANUAL ENTRY ==========
    elif entry_method == "✍️ Manual Entry":
        st.subheader("Manual Deal Entry")
        
        deal_type = st.selectbox("Deal Type:", ["M&A", "Investment"])
        
        if deal_type == "M&A":
            with st.form("ma_form"):
                st.markdown("#### M&A Deal Information")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    company = st.text_input("Company Name*")
                    acquirer = st.text_input("Acquirer*")
                    deal_type_select = st.selectbox("Deal Type*", ["Acquisition", "Merger"])
                
                with col2:
                    deal_value = st.text_input("Deal Value (or 'Undisclosed')*")
                    quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"])
                    month = st.selectbox("Month*", [
                        "January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"
                    ])
                
                technology = st.text_area("Technology/Description*")
                
                submitted = st.form_submit_button("Add M&A Deal")
                
                if submitted:
                    if company and acquirer and technology:
                        new_deal = {
                            'Company': company,
                            'Acquirer': acquirer,
                            'Deal Type (Merger / Acquisition)': deal_type_select,
                            'Technology/Description': technology,
                            'Deal Value': deal_value if deal_value else 'Undisclosed',
                            'Quarter': quarter,
                            'Month': month
                        }
                        
                        if not check_duplicates(new_deal, st.session_state.ma_data, "M&A"):
                            new_row = pd.DataFrame([new_deal])
                            st.session_state.ma_data = pd.concat([st.session_state.ma_data, new_row], ignore_index=True)
                            save_data(st.session_state.ma_data, st.session_state.inv_data)
                            st.success("✅ M&A deal added successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Duplicate deal detected!")
                    else:
                        st.error("Please fill in all required fields (*)")
        
        else:  # Investment
            with st.form("inv_form"):
                st.markdown("#### Investment Deal Information")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    company = st.text_input("Company Name*")
                    funding_type = st.selectbox("Funding Type*", ["VC", "PE"])
                    amount_raised = st.text_input("Amount Raised (or 'Undisclosed')*")
                
                with col2:
                    lead_investors = st.text_input("Lead Investors*")
                    quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"])
                    month = st.selectbox("Month*", [
                        "January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"
                    ])
                
                technology = st.text_area("Technology/Description*")
                
                submitted = st.form_submit_button("Add Investment Deal")
                
                if submitted:
                    if company and lead_investors and technology:
                        new_deal = {
                            'Company': company,
                            'Funding type (VC / PE)': funding_type,
                            'Technology/Description': technology,
                            'Amount Raised': amount_raised if amount_raised else 'Undisclosed',
                            'Lead Investors': lead_investors,
                            'Quarter': quarter,
                            'Month': month
                        }
                        
                        if not check_duplicates(new_deal, st.session_state.inv_data, "Investment"):
                            new_row = pd.DataFrame([new_deal])
                            st.session_state.inv_data = pd.concat([st.session_state.inv_data, new_row], ignore_index=True)
                            save_data(st.session_state.ma_data, st.session_state.inv_data)
                            st.success("✅ Investment deal added successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Duplicate deal detected!")
                    else:
                        st.error("Please fill in all required fields (*)")
    
    # ========== UPLOAD JP MORGAN REPORT ==========
    else:
        st.subheader("Upload JP Morgan Report")
        
        uploaded_file = st.file_uploader("Upload JP Morgan MedTech Report (PDF)", type=['pdf'])
        
        if uploaded_file is not None:
            st.info("PDF parsing functionality - this would extract data from JP Morgan reports")
            st.write(f"File uploaded: {uploaded_file.name}")
            
            if st.button("Process Report"):
                st.warning("PDF processing feature coming soon! This will automatically extract deal data and update the JP Morgan Summary tab.")
                # TODO: Implement PDF parsing logic here
