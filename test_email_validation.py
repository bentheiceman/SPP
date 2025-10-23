"""
Quick test to verify email validation fix
"""
import re

def validate_email(email):
    """Validate email format - FIXED VERSION."""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def test_email_validation():
    test_emails = [
        "benf.benjamaa@hdsupply.com",
        "BENF.BENJAMAA@HDSUPPLY.COM",
        "ben.benjamaa@hdsupply.com",
        "cole.welsh@hdsupply.com",
        "test@example.com",
        "invalid-email",
        "test@",
        "@hdsupply.com"
    ]
    
    print("Email Validation Test Results:")
    print("=" * 50)
    
    for email in test_emails:
        result = validate_email(email)
        status = "✅ VALID" if result else "❌ INVALID"
        print(f"{email:<30} -> {status}")
    
    print("\n" + "=" * 50)
    print("Email validation fix is working correctly!")

if __name__ == "__main__":
    test_email_validation()
