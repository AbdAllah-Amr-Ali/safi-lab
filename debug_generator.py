import sys
import os
import json
import traceback
import time

# Add current directory to path
sys.path.append(os.getcwd())

try:
    from main import SafiLabAPI
except ImportError:
    print("Could not import SafiLabAPI from main.py")
    sys.exit(1)

def debug_generator():
    print("Initializing API...")
    api = SafiLabAPI()
    
    # Create dummy patient
    dummy_patient = {
        "id": "TEST_DATE_GEN",
        "name": "Date Gen Test",
        "age": "30",
        "gender": "Male",
        "date": "2025-12-01"
    }
    
    print(f"Saving dummy patient: {dummy_patient['id']}")
    success = api.save_patient(json.dumps(dummy_patient))
    if not success:
        print("Failed to save dummy patient!")
        return

    print("Patient saved. Fetching details...")
    details_json = api.get_patient_details(dummy_patient['id'])
    details = json.loads(details_json)
    
    print("Last Modified:", details.get('last_modified'))
    
    # Test safe filename logic
    test_name = "John Doe 123"
    safe_name = api._get_safe_filename(test_name)
    print(f"Safe Filename for '{test_name}': '{safe_name}'")
    
    if safe_name == "John Doe 123":
        print("Safe filename logic matches VBA (spaces preserved)!")
    else:
        print(f"Safe filename logic MISMATCH! Expected 'John Doe 123', got '{safe_name}'")

if __name__ == "__main__":
    debug_generator()
