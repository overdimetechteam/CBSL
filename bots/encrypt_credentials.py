#!/usr/bin/env python3
"""
Credential Encryption Utility
Enter the FERNET_KEY, username and password to generate encrypted credentials for .env
"""

from cryptography.fernet import Fernet
from pathlib import Path
root_dir = Path(__file__).parent.parent

print("=" * 60)
print("ERP Credential Encryption Tool")
print("=" * 60)

username = input("\nEnter ERP username: ").strip()
password = input("Enter ERP password: ").strip()

if not username or not password:
    print("❌ Username and password cannot be empty!")
    exit(1)

# Generate new key
key = Fernet.generate_key()
key_str = key.decode()
fernet = Fernet(key)

enc_username = fernet.encrypt(username.encode()).decode()
enc_password = fernet.encrypt(password.encode()).decode()

# Verify
try:
    assert fernet.decrypt(enc_username.encode()).decode() == username
    assert fernet.decrypt(enc_password.encode()).decode() == password
except Exception as e:
    print(f"❌ Verification failed: {e}")
    exit(1)

print("\n" + "=" * 60)
print("SUCCESS")
print("=" * 60)
print(f"\nFERNET_KEY   = {key_str}")
print(f"ENC_USERNAME = {enc_username}")
print(f"ENC_PASSWORD = {enc_password}")
print("\n✅ Verification passed.")
print("⚠️  Update FERNET_KEY in system environment variables and ENC_USERNAME/ENC_PASSWORD in .env")
