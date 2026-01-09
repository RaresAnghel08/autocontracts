import os
from dotenv import load_dotenv
from supabase import create_client, Client

# Load environment variables
load_dotenv()

# Supabase configuration
SUPABASE_URL = os.environ.get('SUPABASE_URL')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY')

def test_connection():
    try:
        print("Testing Supabase connection...")
        print(f"Supabase URL: {SUPABASE_URL}")
        print(f"Supabase Key: {'*' * len(SUPABASE_KEY) if SUPABASE_KEY else 'Not set'}")

        if not SUPABASE_URL or not SUPABASE_KEY:
            print("❌ SUPABASE_URL and SUPABASE_KEY must be set in environment variables")
            return

        # Create Supabase client
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("✅ Supabase client created successfully!")

        # Test a simple query - try to get table info
        try:
            # Try to select from a table that should exist
            response = supabase.table('contract_marshmallow').select('id').limit(1).execute()
            print("✅ Database query successful!")
            print(f"Response data: {response.data}")
        except Exception as e:
            print(f"⚠️  Query test failed (tables may not exist yet): {e}")
            print("This is expected if tables haven't been created yet.")

        print("✅ Supabase connection test completed successfully!")

    except Exception as e:
        print(f"❌ Connection failed: {e}")
        print("\nTroubleshooting steps:")
        print("1. Check if SUPABASE_URL and SUPABASE_KEY are set correctly")
        print("2. Verify the Supabase project is active")
        print("3. Ensure the API key has the correct permissions")
        print("4. Check network connectivity")

if __name__ == '__main__':
    test_connection()