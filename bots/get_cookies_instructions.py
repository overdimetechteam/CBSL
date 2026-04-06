#!/usr/bin/env python3
"""
Instructions for getting session cookies manually
"""

import json
import os

def print_instructions():
    """Print step-by-step instructions for getting session cookies"""
    
    instructions = """
=== MANUAL SESSION COOKIES SETUP ===

Since we don't have Selenium installed, you need to manually get session cookies from your browser.

STEP 1: Login to the website manually
1. Open your browser (Chrome, Firefox, Edge)
2. Go to: https://erp.assetline.lk/
3. Login with your credentials
4. Navigate to: https://erp.assetline.lk/Application/Home/RECOVERY
5. Click on "Information Center" tile
6. Click on "Contract Inquiry" tile
7. Click on "Add" button
8. Enter any contract number (e.g., MUUV241723430) and press Enter
9. Wait for the search to complete

STEP 2: Get cookies from browser
1. Open Developer Tools (F12 or right-click -> Inspect)
2. Go to Application/Storage tab (Chrome) or Storage tab (Firefox)
3. Look for "Cookies" in the left sidebar
4. Click on "https://erp.assetline.lk"
5. You should see all cookies for the domain

STEP 3: Copy cookies to session_cookies.json
Create a file called 'session_cookies.json' in the bots directory with this format:

[
    {
        "name": "cookie_name_1",
        "value": "cookie_value_1",
        "domain": ".erp.assetline.lk",
        "path": "/"
    },
    {
        "name": "cookie_name_2", 
        "value": "cookie_value_2",
        "domain": ".erp.assetline.lk",
        "path": "/"
    }
]

IMPORTANT COOKIES TO INCLUDE:
- Any session cookies (usually contain "session", "auth", "token" in the name)
- ASP.NET session cookies
- Authentication cookies

STEP 4: Run the bot
Once you have session_cookies.json file, run:
python na_contract_numbers_search_bot_manual.py

TROUBLESHOOTING:
- If you get "LOGOUT" response, your session has expired
- If you get "INVALID REQUEST", you may need to include a verification token
- Make sure you're logged in and have accessed the Contract Inquiry page before getting cookies

=== END INSTRUCTIONS ===
"""
    
    print(instructions)

def create_sample_cookies_file():
    """Create a sample session_cookies.json file"""
    
    sample_cookies = [
        {
            "name": "ASP.NET_SessionId",
            "value": "YOUR_SESSION_ID_HERE",
            "domain": ".erp.assetline.lk",
            "path": "/"
        },
        {
            "name": "AuthToken",
            "value": "YOUR_AUTH_TOKEN_HERE", 
            "domain": ".erp.assetline.lk",
            "path": "/"
        }
    ]
    
    filename = "session_cookies.json"
    
    if not os.path.exists(filename):
        with open(filename, 'w') as f:
            json.dump(sample_cookies, f, indent=2)
        print(f"Created sample {filename} file")
        print("Please edit this file with your actual cookie values")
    else:
        print(f"{filename} already exists")

if __name__ == "__main__":
    print_instructions()
    print("\n" + "="*50 + "\n")
    create_sample_cookies_file()
