"""
Test Snowflake Connection
Simple script to test the Snowflake connection with the new configuration.
"""

from spp_metric_automation import SPPMetricAutomation
import logging

def test_connection():
    """Test the Snowflake connection."""
    print("Testing Snowflake Connection...")
    print("=" * 40)
    
    try:
        # Create automation instance
        automation = SPPMetricAutomation()
        
        # Test connection
        if automation.connect_to_snowflake():
            print("✅ Connection successful!")
            
            # Test a simple query
            print("\nTesting simple query...")
            if automation.connection:
                cursor = automation.connection.cursor()
                cursor.execute("SELECT CURRENT_USER(), CURRENT_ACCOUNT(), CURRENT_TIMESTAMP()")
                result = cursor.fetchone()
                
                print(f"Current User: {result[0]}")
                print(f"Current Account: {result[1]}")
                print(f"Current Timestamp: {result[2]}")
                
                cursor.close()
                automation.connection.close()
            
            print("\n✅ Query test successful!")
            return True
            
        else:
            print("❌ Connection failed!")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    # Set logging level to see connection details
    logging.basicConfig(level=logging.INFO)
    
    success = test_connection()
    
    if success:
        print("\n🎉 Snowflake connection is working correctly!")
        print("You can now run the full automation.")
    else:
        print("\n⚠️  Please check your configuration and try again.")
        print("Make sure you have access to Snowflake and the required permissions.")
