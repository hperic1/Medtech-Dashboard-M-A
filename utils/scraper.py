import re
import requests
from bs4 import BeautifulSoup
import streamlit as st

def scrape_article(url):
    """Scrape article content from URL and extract deal information"""
    try:
        # Fetch the article
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # Parse HTML
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Extract text content
        article_text = soup.get_text(separator=' ', strip=True)
        
        # Parse for deals
        deals = parse_deal_from_text(article_text)
        
        return deals
    
    except Exception as e:
        st.error(f"Error fetching article: {str(e)}")
        return []

def parse_deal_from_text(text):
    """Extract deal information from text using pattern matching"""
    deals = []
    
    # Common patterns for M&A deals
    ma_patterns = [
        r'(\w+(?:\s+\w+)*)\s+(?:acquired|acquires|to acquire)\s+(\w+(?:\s+\w+)*)\s+for\s+\$?([\d.]+)\s*(billion|million|B|M)',
        r'(\w+(?:\s+\w+)*)\s+(?:acquired|acquires|to acquire)\s+(\w+(?:\s+\w+)*)',
        r'(\w+(?:\s+\w+)*)\s+and\s+(\w+(?:\s+\w+)*)\s+(?:merge|merged|to merge)',
    ]
    
    # Common patterns for investment deals
    inv_patterns = [
        r'(\w+(?:\s+\w+)*)\s+(?:raised|raises|secures)\s+\$?([\d.]+)\s*(billion|million|B|M)\s+in\s+Series\s+([A-Z])',
        r'(\w+(?:\s+\w+)*)\s+(?:raised|raises|secures)\s+\$?([\d.]+)\s*(billion|million|B|M)',
    ]
    
    # Try M&A patterns
    for pattern in ma_patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            groups = match.groups()
            
            deal = {
                'Company': groups[1] if len(groups) > 1 else 'Unknown',
                'Acquirer': groups[0] if len(groups) > 0 else 'Unknown',
                'Deal Type (Merger / Acquisition)': 'Merger' if 'merge' in match.group(0).lower() else 'Acquisition',
                'Technology/Description': 'Extracted from article - needs verification',
                'Deal Value': format_deal_value(groups[2], groups[3]) if len(groups) > 3 else 'Undisclosed',
                'Quarter': 'Undisclosed',
                'Month': 'Undisclosed'
            }
            
            deals.append(deal)
    
    # Try investment patterns
    for pattern in inv_patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            groups = match.groups()
            
            deal = {
                'Company': groups[0] if len(groups) > 0 else 'Unknown',
                'Funding type (VC / PE)': 'VC',
                'Technology/Description': 'Extracted from article - needs verification',
                'Amount Raised': format_deal_value(groups[1], groups[2]) if len(groups) > 2 else 'Undisclosed',
                'Lead Investors': 'Unknown - verify from article',
                'Quarter': 'Undisclosed',
                'Month': 'Undisclosed'
            }
            
            deals.append(deal)
    
    # Remove duplicates
    unique_deals = []
    seen = set()
    for deal in deals:
        # Create a key for deduplication
        if 'Company' in deal and 'Acquirer' in deal:
            key = f"{deal['Company']}_{deal['Acquirer']}"
        elif 'Company' in deal:
            key = f"{deal['Company']}_{deal.get('Amount Raised', '')}"
        else:
            continue
        
        if key not in seen:
            seen.add(key)
            unique_deals.append(deal)
    
    return unique_deals

def format_deal_value(amount, unit):
    """Format deal value to standard format"""
    try:
        amount_float = float(amount)
        
        # Normalize unit
        unit_lower = unit.lower()
        if unit_lower in ['billion', 'b']:
            return f"${amount_float}B"
        elif unit_lower in ['million', 'm']:
            return f"${amount_float}M"
        else:
            return f"${amount_float}"
    except:
        return 'Undisclosed'

def extract_deal_date(text):
    """Extract date from text if available"""
    # Common date patterns
    date_patterns = [
        r'(\w+)\s+(\d{1,2}),?\s+(\d{4})',  # January 15, 2025
        r'(\d{1,2})/(\d{1,2})/(\d{4})',     # 01/15/2025
        r'(\d{4})-(\d{2})-(\d{2})',          # 2025-01-15
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            # Return the matched date - further parsing can be added
            return match.group(0)
    
    return None

def extract_technology_description(text, company_name):
    """Extract technology description near company mention"""
    # Find sentences containing the company name
    sentences = text.split('.')
    
    for sentence in sentences:
        if company_name.lower() in sentence.lower():
            # Clean and return the sentence
            cleaned = sentence.strip()
            if len(cleaned) > 20:  # Minimum length check
                return cleaned
    
    return 'Technology description not found - please add manually'
