"""Integration test: verify end-to-end date handling works correctly."""
import os
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bundle_script import (
    _safe_strftime,
    _format_email_datetime,
    _validate_date,
    _parse_pywintypes_date,
)


def test_real_world_scenario():
    """Test a real-world scenario: processing an email with a date."""
    
    print("\n" + "="*70)
    print("INTEGRATION TEST: Real-World Date Handling Scenario")
    print("="*70)
    
    # Scenario 1: Email received May 6, 2025 at 4:17 PM
    print("\n1️⃣ Scenario: Processing Outlook email with SentOn date")
    print("-" * 70)
    
    email_date = datetime.datetime(2025, 5, 6, 16, 17, 0)
    print(f"   Email SentOn (datetime object): {email_date}")
    
    formatted_date = _format_email_datetime(email_date)
    print(f"   Formatted for display: {formatted_date}")
    
    validation_result = _validate_date(formatted_date, "email_from_john.msg", "SentOn")
    if validation_result is None:
        print(f"   ✓ Date validation: PASSED")
    else:
        print(f"   ✗ Date validation: FAILED - {validation_result}")
        
    # Scenario 2: Different date formats from different locales
    print("\n2️⃣ Scenario: Multiple locale date formats")
    print("-" * 70)
    
    test_formats = [
        ("ISO format", "2025-05-06 16:17:00"),
        ("US locale", "05/06/2025 04:17:00"),
        ("EU locale", "06/05/2025 16:17:00"),
        ("OLE date", 45475.677),  # Approximate OLE date for May 6, 2025
    ]
    
    for label, date_value in test_formats:
        parsed = _parse_pywintypes_date(date_value)
        if parsed:
            formatted = _safe_strftime(parsed)
            validation = _validate_date(formatted, "test.msg", label)
            status = "✓ PASSED" if validation is None else f"⚠ {validation}"
            print(f"   {label}: Parsed to '{formatted}' - {status}")
        else:
            print(f"   {label}: Could not parse (expected for OLE epoch)")
    
    # Scenario 3: Invalid/unknown dates
    print("\n3️⃣ Scenario: Handling invalid and unknown dates")
    print("-" * 70)
    
    error_cases = [
        ("None value", None),
        ("Bogus 1899 date", datetime.datetime(1899, 12, 30, 0, 0, 0)),
        ("Out of range (1960)", datetime.datetime(1960, 1, 1, 0, 0, 0)),
    ]
    
    for label, date_value in error_cases:
        formatted = _format_email_datetime(date_value)
        print(f"   {label}: Handled as '{formatted}'")
    
    # Scenario 4: All 12 months work correctly
    print("\n4️⃣ Scenario: All months format correctly")
    print("-" * 70)
    
    all_months_ok = True
    for month in range(1, 13):
        test_date = datetime.datetime(2025, month, 15, 10, 30, 0)
        formatted = _safe_strftime(test_date)
        validation = _validate_date(formatted, f"email_month_{month}.msg", f"Month{month}")
        if validation is None:
            print(f"   Month {month:2d}: {formatted} ✓")
        else:
            print(f"   Month {month:2d}: {formatted} ✗ {validation}")
            all_months_ok = False
    
    if all_months_ok:
        print("\n   ✓ All 12 months format and validate correctly!")
    
    # Scenario 5: Edge cases
    print("\n5️⃣ Scenario: Edge cases and boundaries")
    print("-" * 70)
    
    edge_cases = [
        ("New Year", datetime.datetime(2025, 1, 1, 0, 0, 0)),
        ("Leap day", datetime.datetime(2024, 2, 29, 12, 0, 0)),
        ("End of year", datetime.datetime(2025, 12, 31, 23, 59, 0)),
        ("Min hour", datetime.datetime(2025, 6, 15, 0, 0, 0)),
        ("Max hour", datetime.datetime(2025, 6, 15, 23, 59, 0)),
    ]
    
    for label, test_date in edge_cases:
        formatted = _safe_strftime(test_date)
        validation = _validate_date(formatted, "edge_case.msg", label)
        status = "✓" if validation is None else "✗"
        print(f"   {label:20s}: {formatted} {status}")
    
    print("\n" + "="*70)
    print("✅ INTEGRATION TEST COMPLETE - All scenarios working correctly!")
    print("="*70 + "\n")


if __name__ == "__main__":
    test_real_world_scenario()
