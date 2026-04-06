# Security Documentation - NA Contract Numbers Search Bot

## 🔒 Security Overview

This bot implements multiple layers of security to protect sensitive contract data and credentials.

## 🛡️ Security Features

### 1. **Credential Encryption**
- ✅ **Fernet Symmetric Encryption**: Username and password are encrypted using `cryptography.fernet`
- ✅ **Separate Key File**: Encryption key stored separately from encrypted data
- ✅ **Environment Variables**: Encrypted credentials stored in `.env` file

### 2. **Secure Logging**
- ✅ **Configurable Logging**: `SECURE_LOGGING = True` prevents sensitive data in logs
- ✅ **Token Masking**: Verification tokens and cookies are not logged in production
- ✅ **Data Masking**: Contract data is not logged in production mode

### 3. **File Encryption**
- ✅ **Output Encryption**: All output files can be encrypted using `ENCRYPT_OUTPUT_FILES = True`
- ✅ **Separate Key Storage**: Each encrypted file has its own key file
- ✅ **Binary Storage**: Encrypted files stored as binary to prevent text extraction

### 4. **Session Security**
- ✅ **HTTPS Only**: All API calls use HTTPS
- ✅ **Session Cookies**: Proper session management with cookies
- ✅ **Verification Tokens**: ASP.NET anti-forgery tokens included in requests

### 5. **Data Cleanup**
- ✅ **Automatic Cleanup**: `CLEANUP_SENSITIVE_DATA = True` removes temporary files
- ✅ **Secure Deletion**: Temporary files are properly removed after processing

## ⚙️ Security Configuration

### Security Settings
```python
SECURE_LOGGING = True              # Prevents sensitive data logging
ENCRYPT_OUTPUT_FILES = True        # Encrypts all output files
CLEANUP_SENSITIVE_DATA = True      # Cleans up temporary files
```

### File Structure
```
Script/
├── .key                          # Encryption key (KEEP SECURE!)
├── .env                          # Encrypted credentials
├── bots/
│   ├── na_contract_numbers_search_bot_api.py
│   └── outputs/
│       └── contract_search_results/
│           ├── *.json            # Encrypted output files
│           └── *.json.key        # Encryption keys for output files
```

## 🚨 Security Best Practices

### 1. **File Permissions**
```bash
# Set restrictive permissions on sensitive files
chmod 600 .key
chmod 600 .env
chmod 700 bots/outputs/
```

### 2. **Environment Security**
- ✅ Never commit `.key` or `.env` files to version control
- ✅ Use `.gitignore` to exclude sensitive files
- ✅ Store encryption keys in secure key management systems (production)

### 3. **Network Security**
- ✅ All communications use HTTPS
- ✅ Session cookies are properly managed
- ✅ Verification tokens prevent CSRF attacks

### 4. **Data Protection**
- ✅ Sensitive data is encrypted at rest
- ✅ Temporary files are automatically cleaned up
- ✅ Logs don't contain sensitive information in production

## 🔍 Security Validation

The bot includes automatic security validation:

```python
def validate_security_settings():
    # Checks:
    # - Secure logging enabled
    # - File encryption enabled
    # - Cleanup enabled
    # - Required security files exist
    # - File permissions are appropriate
```

## 🚨 Security Warnings

### ⚠️ **Critical Security Notes:**

1. **Key Management**: The `.key` file contains the master encryption key. Protect it at all costs!
2. **File Permissions**: Ensure sensitive files have restrictive permissions
3. **Network Security**: Only run on secure, trusted networks
4. **Access Control**: Limit access to the bot and its output files
5. **Audit Logging**: Monitor access to sensitive files and data

### 🔐 **Production Security Checklist:**

- [ ] Encryption key stored in secure key management system
- [ ] File permissions set to 600 for sensitive files
- [ ] Network access restricted to authorized users
- [ ] Regular security audits performed
- [ ] Backup encryption keys securely
- [ ] Monitor for unauthorized access attempts

## 🛠️ Security Troubleshooting

### Common Security Issues:

1. **"Encryption key file not found"**
   - Ensure `.key` file exists in the root directory
   - Check file permissions (should be 600)

2. **"Environment file not found"**
   - Ensure `.env` file exists with encrypted credentials
   - Run the encryption setup script if needed

3. **"File permissions too permissive"**
   - Set restrictive permissions: `chmod 600 .env .key`

4. **"Security validation failed"**
   - Review security settings in the bot configuration
   - Ensure all security features are enabled

## 📞 Security Contact

For security issues or questions:
- Review this documentation
- Check the security validation output
- Ensure all security settings are properly configured

---

**⚠️ IMPORTANT**: This bot handles sensitive contract data. Always follow security best practices and regularly audit your security configuration.
