import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="MedTech M&A Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to make tables span full width
st.markdown("""
    <style>
    .dataframe {
        width: 100% !important;
    }
    .stDataFrame {
        width: 100%;
    }
    div[data-testid="stDataFrame"] > div {
        width: 100%;
    }
    .main .block-container {
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 100%;
    }
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def load_data():
    """Load data from Excel file"""
    try:
        excel_file = '/mnt/project/MedTech_MA_Masterlist.xlsx'
        
        # Load all sheets
        h1_ma = pd.read_excel(excel_file, sheet_name='H1 M&A Activity')
        h1_investment = pd.read_excel(excel_file, sheet_name='H1 Investment_Activity')
        h2_ma = pd.read_excel(excel_file, sheet_name='H2 M&A Activity')
        h2_investment = pd.read_excel(excel_file, sheet_name='H2 Investment_Activity')
        
        return {
            'h1_ma': h1_ma,
            'h1_investment': h1_investment,
            'h2_ma': h2_ma,
            'h2_investment': h2_investment
        }
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def create_volume_chart(data_dict):
    """Create M&A volume chart"""
    h1_count = len(data_dict['h1_ma'])
    h2_count = len(data_dict['h2_ma'])
    
    fig = go.Figure(data=[
        go.Bar(
            x=['H1', 'H2'],
            y=[h1_count, h2_count],
            text=[h1_count, h2_count],
            textposition='auto',
            marker_color=['#1f77b4', '#ff7f0e']
        )
    ])
    
    fig.update_layout(
        title="M&A Activity Volume by Half",
        xaxis_title="Period",
        yaxis_title="Number of Deals",
        height=500,
        margin=dict(t=50, b=50, l=50, r=50),
        showlegend=False
    )
    
    return fig

def create_investment_chart(data_dict):
    """Create investment activity chart"""
    h1_count = len(data_dict['h1_investment'])
    h2_count = len(data_dict['h2_investment'])
    
    fig = go.Figure(data=[
        go.Bar(
            x=['H1', 'H2'],
            y=[h1_count, h2_count],
            text=[h1_count, h2_count],
            textposition='auto',
            marker_color=['#2ca02c', '#d62728']
        )
    ])
    
    fig.update_layout(
        title="Investment Activity Volume by Half",
        xaxis_title="Period",
        yaxis_title="Number of Investments",
        height=500,
        margin=dict(t=50, b=50, l=50, r=50),
        showlegend=False
    )
    
    return fig

def main():
    st.title("🏥 MedTech M&A Activity Dashboard")
    st.markdown("---")
    
    # Load data
    data = load_data()
    
    if data is None:
        st.error("Failed to load data. Please check the file path and format.")
        return
    
    # Sidebar filters
    st.sidebar.header("Filters")
    view_option = st.sidebar.selectbox(
        "Select View",
        ["Overview", "H1 M&A Activity", "H1 Investment Activity", 
         "H2 M&A Activity", "H2 Investment Activity"]
    )
    
    # Main content area
    if view_option == "Overview":
        st.header("Overview")
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("H1 M&A Deals", len(data['h1_ma']))
        with col2:
            st.metric("H2 M&A Deals", len(data['h2_ma']))
        with col3:
            st.metric("H1 Investments", len(data['h1_investment']))
        with col4:
            st.metric("H2 Investments", len(data['h2_investment']))
        
        st.markdown("---")
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            st.plotly_chart(create_volume_chart(data), use_container_width=True)
        
        with col2:
            st.plotly_chart(create_investment_chart(data), use_container_width=True)
    
    elif view_option == "H1 M&A Activity":
        st.header("H1 M&A Activity")
        st.metric("Total Deals", len(data['h1_ma']))
        st.dataframe(data['h1_ma'], use_container_width=True, height=600)
    
    elif view_option == "H1 Investment Activity":
        st.header("H1 Investment Activity")
        st.metric("Total Investments", len(data['h1_investment']))
        st.dataframe(data['h1_investment'], use_container_width=True, height=600)
    
    elif view_option == "H2 M&A Activity":
        st.header("H2 M&A Activity")
        st.metric("Total Deals", len(data['h2_ma']))
        st.dataframe(data['h2_ma'], use_container_width=True, height=600)
    
    elif view_option == "H2 Investment Activity":
        st.header("H2 Investment Activity")
        st.metric("Total Investments", len(data['h2_investment']))
        st.dataframe(data['h2_investment'], use_container_width=True, height=600)
    
    # Footer
    st.markdown("---")
    st.markdown(f"*Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")

if __name__ == "__main__":
    main()