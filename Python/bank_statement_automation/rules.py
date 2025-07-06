"""
Categorization rules for bank statement transactions
"""

import pandas as pd

# Dictionary of categories and their associated keywords
CATEGORY_RULES = {
    "Food & Dining": [
        "starbucks", "mcdonalds", "burger", "pizza", "restaurant", "cafe", 
        "coffee", "food", "dining", "subway", "chipotle", "taco", "sushi",
        "grubhub", "doordash", "ubereats", "delivery", "takeout"
    ],
    
    "Income": [
        "payroll", "salary", "deposit", "income", "payment received", 
        "direct deposit", "paycheck", "bonus", "commission", "refund"
    ],
    
    "Housing": [
        "rent", "mortgage", "housing", "apartment", "home", "lease",
        "property", "real estate", "landlord"
    ],
    
    "Shopping": [
        "amazon", "walmart", "target", "costco", "shopping", "retail",
        "best buy", "home depot", "lowes", "ikea", "clothing", "apparel"
    ],
    
    "Transportation": [
        "uber", "lyft", "gas", "fuel", "transportation", "parking",
        "taxi", "public transit", "metro", "bus", "train", "car"
    ],
    
    "Utilities": [
        "electric", "water", "gas", "internet", "phone", "utility",
        "electricity", "cable", "wifi", "cellular", "mobile"
    ],
    
    "Entertainment": [
        "netflix", "spotify", "movie", "theater", "entertainment", "gym",
        "hulu", "disney", "youtube", "music", "games", "sports"
    ],
    
    "Healthcare": [
        "doctor", "pharmacy", "medical", "health", "dental", "hospital",
        "clinic", "insurance", "prescription", "therapy"
    ],
    
    "Education": [
        "tuition", "school", "college", "university", "education", "books",
        "course", "training", "workshop", "seminar"
    ],
    
    "Travel": [
        "hotel", "airline", "flight", "travel", "vacation", "booking",
        "airbnb", "expedia", "booking.com", "trip"
    ],
    
    "Insurance": [
        "insurance", "auto insurance", "home insurance", "life insurance",
        "health insurance", "premium"
    ],
    
    "Investments": [
        "investment", "stock", "bond", "mutual fund", "etf", "brokerage",
        "fidelity", "vanguard", "schwab", "robinhood"
    ]
}

def categorize_transaction(description):
    """
    Categorize a transaction based on its description
    
    Args:
        description (str): Transaction description
        
    Returns:
        str: Category name
    """
    if not description or pd.isna(description):
        return "Uncategorized"
    
    desc = str(description).lower()
    
    # Check each category's keywords
    for category, keywords in CATEGORY_RULES.items():
        if any(keyword in desc for keyword in keywords):
            return category
    
    return "Uncategorized"

def add_custom_rule(category, keywords):
    """
    Add custom categorization rules
    
    Args:
        category (str): Category name
        keywords (list): List of keywords to match
    """
    if category not in CATEGORY_RULES:
        CATEGORY_RULES[category] = []
    
    CATEGORY_RULES[category].extend(keywords)
    print(f"✅ Added {len(keywords)} keywords to '{category}' category")

def get_category_summary():
    """
    Get a summary of all categories and their keyword counts
    
    Returns:
        dict: Category names and their keyword counts
    """
    return {category: len(keywords) for category, keywords in CATEGORY_RULES.items()} 