# NA Contract Numbers Search Bot

This bot automates the process of searching for contract numbers in the ERP Assetline system with **encrypted credential storage**.

## Features

- **🔐 Encrypted Authentication**: Credentials are encrypted using Fernet encryption
- **Secure Storage**: Uses separate `.key` and `.env` files for maximum security
- **Headless Automation**: Runs Chrome in headless mode using Selenium WebDriver
- **Network Monitoring**: Captures API responses and network traffic
- **Data Export**: Saves results in both JSON and CSV formats
- **Comprehensive Logging**: Detailed logging for debugging and monitoring
- **Error Handling**: Robust error handling with retry mechanisms

## Prerequisites

- Python 3.8 or higher
- Chrome browser installed
- Access to the ERP Assetline system

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set up encrypted credentials:**
   ```bash
   cd bots
   python encrypt_credentials.py
   ```
   
   Follow the interactive menu to:
   - Generate an encryption key
   - Encrypt your ERP credentials
   - Test the decryption

3. **Ensure Chrome is installed:**
   The bot uses Selenium WebDriver which requires Chrome browser to be installed.

## 🔐 Security Setup

### Step 1: Generate Encryption Key
```bash
cd bots
python encrypt_credentials.py
# Choose option 1: Generate new encryption key
```

This creates a `.key` file in your root directory. **Keep this file secure and never share it!**

### Step 2: Encrypt Your Credentials
```bash
# Choose option 2: Encrypt credentials
# Enter your ERP username and password when prompted
```

This creates a `.env` file with encrypted credentials.

### Step 3: Verify Setup
```bash
# Choose option 3: Test decryption
# This verifies your encrypted credentials work correctly
```

## File Structure

```
your_project/
├── .key                    # 🔑 Encryption key (KEEP SECURE!)
├── .env                    # 🔒 Encrypted credentials
├── requirements.txt        # Python dependencies
├── bots/
│   ├── na_contract_numbers_search_bot.py
│   ├── encrypt_credentials.py
│   └── README_NA_Contract_Bot.md
├── outputs/
│   └── contract_search_results/
└── logs/
    └── contract_search_bot.log
```

## Configuration

### Environment Variables (Automatically Created)

- `ENC_USERNAME`: Your encrypted ERP username
- `ENC_PASSWORD`: Your encrypted ERP password

### Contract Numbers

The bot currently uses hardcoded contract numbers:
```python
na_contract_numbers_data = [
    "MUUV241723430",
    "MUUV241766230", 
    "MUUV241780180",
    "MUUV251822000",
    "MUUV251835020",
    "MUUV251857820",
    "MWAL250127060",
    "MWAT250036110"
]
```

To use different contract numbers, modify the `na_contract_numbers_data` list in the script.

## Usage

### Basic Usage

```bash
cd bots
python na_contract_numbers_search_bot.py
```

### What the Bot Does

1. **🔓 Decrypt**: Automatically decrypts your ERP credentials using the `.key` file
2. **🔐 Login**: Authenticates with the ERP system using decrypted credentials
3. **🧭 Navigation**: Navigates through the system to reach the Contract Inquiry page
4. **🔍 Search**: For each contract number:
   - Enters the contract number in the search field
   - Presses Enter to search
   - Captures the API response
   - Saves results to files
5. **📁 Output**: Saves results in `outputs/contract_search_results/` directory

### Output Files

- **JSON Files**: `{contract_number}_response.json` - Contains full API response
- **CSV Files**: `{contract_number}_results.csv` - Flattened data for analysis
- **Logs**: `logs/contract_search_bot.log` - Detailed execution logs

## Security Features

- **🔐 Fernet Encryption**: Industry-standard symmetric encryption
- **🔑 Separate Key File**: Encryption key stored separately from encrypted data
- **🔒 Environment Variables**: Encrypted credentials loaded securely
- **Headless Mode**: Chrome runs without visible UI
- **Secure Logging**: No credential information is ever logged
- **Error Handling**: Graceful failure without exposing sensitive data

## 🔑 Managing Your Encryption Key

### Backup Your Key
- Store the `.key` file in a secure location
- Consider using your system's key vault or password manager
- Never commit the `.key` file to version control

### If You Lose Your Key
- You'll need to regenerate a new key
- Re-encrypt credentials with the new key
- Old encrypted data cannot be recovered without the original key

### Key Rotation
- Generate a new key periodically for enhanced security
- Re-encrypt credentials with the new key
- Delete the old `.key` file

## Troubleshooting

### Common Issues

1. **Missing .key file:**
   ```
   ❌ Encryption key file not found: .key
   ```
   **Solution**: Run `python encrypt_credentials.py` and generate a new key

2. **Missing .env file:**
   ```
   ❌ Environment file not found: .env
   ```
   **Solution**: Run `python encrypt_credentials.py` and encrypt your credentials

3. **Decryption failed:**
   ```
   ❌ Failed to decrypt credentials: InvalidToken
   ```
   **Solution**: Your `.key` and `.env` files don't match. Regenerate both.

4. **ChromeDriver Issues:**
   - The bot now uses `webdriver-manager` to automatically download ChromeDriver
   - Ensure Chrome browser is installed and up to date
   - If you get ChromeDriver errors, try updating Chrome browser

5. **ModuleNotFoundError: No module named 'distutils':**
   ```
   ModuleNotFoundError: No module named 'distutils'
   ```
   **Solution**: Install setuptools: `pip install setuptools>=65.0.0`

6. **Login Failures:**
   - Verify your credentials are correct
   - Check if ERP system is accessible
   - Ensure account is not locked

### Debug Mode

To run with visible Chrome (for debugging):
```python
# Comment out this line in start_driver():
# options.add_argument("--headless=new")
```

## Customization

### Adding New Contract Numbers

Modify the `na_contract_numbers_data` list:
```python
na_contract_numbers_data = [
    "NEW_CONTRACT_001",
    "NEW_CONTRACT_002",
    # ... your contract numbers
]
```

### Changing Output Directory

Modify the `OUT_DIR` variable:
```python
OUT_DIR = "your/custom/path"
```

### Adjusting Timeouts

Modify timeout values in various functions:
```python
def search_contract(driver, contract_no, timeout=20):  # Increased from 15
```

## Logging

The bot provides comprehensive logging:

- **INFO**: Normal operations and successful actions
- **WARNING**: Non-critical issues (e.g., no API response)
- **ERROR**: Critical failures that prevent operation

Logs are saved to both console and file for easy debugging.

## Performance

- **Headless Mode**: Faster execution without UI rendering
- **Network Monitoring**: Efficient capture of API responses
- **Batch Processing**: Processes multiple contracts sequentially
- **Delay Between Searches**: 2-second delay to avoid overwhelming the server

## Support

For issues or questions:
1. Check the logs in `logs/contract_search_bot.log`
2. Verify your `.key` and `.env` files exist and are correct
3. Test decryption: `python encrypt_credentials.py` → option 3
4. Ensure Chrome browser is properly installed
5. Check ERP system accessibility

## Security Best Practices

1. **Never commit `.key` or `.env` files to version control**
2. **Store the `.key` file in a secure location**
3. **Use strong, unique passwords for your ERP account**
4. **Rotate encryption keys periodically**
5. **Limit access to the `.key` file**
6. **Monitor logs for any suspicious activity**

## License

This bot is provided as-is for internal use. Ensure compliance with your organization's policies and the ERP system's terms of service.
