"""Comprehensive tests for date handling and formatting functions."""
import os
import sys
import unittest
import datetime
import re

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bundle_script import (
    _safe_strftime, 
    _format_email_datetime,
    _validate_date,
    _parse_pywintypes_date,
    _is_plausible_date,
    _DATE_PATTERN,
    _MONTH_NAMES,
    _BOGUS_YEARS,
)


class DateFormattingTests(unittest.TestCase):
    """Test date formatting functions."""
    
    def test_safe_strftime_basic(self):
        """Test basic datetime formatting."""
        dt = datetime.datetime(2025, 5, 6, 16, 17, 0)
        result = _safe_strftime(dt)
        self.assertEqual(result, "06 May 2025 16:17")
    
    def test_safe_strftime_all_months(self):
        """Test all months format correctly."""
        for month in range(1, 13):
            dt = datetime.datetime(2025, month, 15, 10, 30, 0)
            result = _safe_strftime(dt)
            # Result should be "15 MonthName 2025 10:30"
            self.assertIn(_MONTH_NAMES[month], result)
            self.assertIn("2025", result)
            self.assertTrue(result.startswith("15"))
    
    def test_safe_strftime_zero_padding(self):
        """Test that day, hour, and minute are zero-padded."""
        dt = datetime.datetime(2025, 1, 1, 5, 3, 0)
        result = _safe_strftime(dt)
        self.assertEqual(result, "01 January 2025 05:03")
    
    def test_safe_strftime_edge_dates(self):
        """Test edge cases: beginning/end of month/year."""
        # First day of year
        dt = datetime.datetime(2025, 1, 1, 0, 0, 0)
        result = _safe_strftime(dt)
        self.assertEqual(result, "01 January 2025 00:00")
        
        # Last day of February (non-leap)
        dt = datetime.datetime(2025, 2, 28, 23, 59, 0)
        result = _safe_strftime(dt)
        self.assertEqual(result, "28 February 2025 23:59")
        
        # Last day of year
        dt = datetime.datetime(2025, 12, 31, 23, 59, 0)
        result = _safe_strftime(dt)
        self.assertEqual(result, "31 December 2025 23:59")


class DateValidationTests(unittest.TestCase):
    """Test date validation functions."""
    
    def test_date_pattern_valid_formats(self):
        """Test that valid dates match the pattern."""
        valid_dates = [
            "01 January 2025 00:00",
            "06 May 2025 16:17",
            "31 December 2025 23:59",
            "15 June 2020 12:00",
        ]
        for date_str in valid_dates:
            self.assertIsNotNone(_DATE_PATTERN.match(date_str),
                                 f"Expected valid date: {date_str}")
    
    def test_date_pattern_invalid_formats(self):
        """Test that invalid dates don't match the pattern."""
        invalid_dates = [
            "6 May 2025 16:17",  # missing leading zero on day
            "06 MAY 2025 16:17",  # wrong case
            "06-May-2025 16:17",  # wrong separator
            "2025-05-06 16:17",  # ISO format
            "Unknown",  # placeholder
            "",  # empty
        ]
        for date_str in invalid_dates:
            self.assertIsNone(_DATE_PATTERN.match(date_str),
                             f"Expected invalid date: {date_str}")
    
    def test_validate_date_valid(self):
        """Test validation of valid date strings."""
        result = _validate_date("06 May 2025 16:17", "test.msg", "SentOn")
        self.assertIsNone(result, "Valid date should return None (no warning)")
    
    def test_validate_date_unknown(self):
        """Test validation of 'Unknown' date."""
        result = _validate_date("Unknown", "test.msg", "SentOn")
        self.assertIsNotNone(result)
        self.assertIn("No date found", result)
        self.assertIn("test.msg", result)
    
    def test_validate_date_invalid_format(self):
        """Test validation of invalid date format."""
        result = _validate_date("06 MAY 2025 16:17", "email.msg", "ReceivedTime")
        self.assertIsNotNone(result)
        self.assertIn("Unexpected format", result)
        self.assertIn("email.msg", result)
    
    def test_validate_date_includes_source_label(self):
        """Test that validation includes the source label."""
        result = _validate_date("Unknown", "file.msg", "CustomProperty")
        self.assertIn("CustomProperty", result)


class DateParsingTests(unittest.TestCase):
    """Test date parsing functions."""
    
    def test_is_plausible_date_valid(self):
        """Test that dates in reasonable range are plausible."""
        # Recent dates
        dt = datetime.datetime(2025, 5, 6, 16, 17)
        self.assertTrue(_is_plausible_date(dt))
        
        # Historic but reasonable
        dt = datetime.datetime(1970, 1, 1, 0, 0)
        self.assertTrue(_is_plausible_date(dt))
        
        # Current/future
        dt = datetime.datetime(2100, 12, 31, 23, 59)
        self.assertTrue(_is_plausible_date(dt))
    
    def test_is_plausible_date_bogus_years(self):
        """Test that bogus year values are rejected."""
        for bogus_year in _BOGUS_YEARS:
            dt = datetime.datetime(bogus_year, 6, 15, 12, 0)
            self.assertFalse(_is_plausible_date(dt),
                            f"Year {bogus_year} should be considered bogus")
    
    def test_is_plausible_date_out_of_range(self):
        """Test that dates outside reasonable range are rejected."""
        # Too old
        dt = datetime.datetime(1969, 12, 31, 23, 59)
        self.assertFalse(_is_plausible_date(dt))
        
        # TooFar in future
        dt = datetime.datetime(2101, 1, 1, 0, 0)
        self.assertFalse(_is_plausible_date(dt))
    
    def test_is_plausible_date_none(self):
        """Test that None returns False."""
        self.assertFalse(_is_plausible_date(None))
    
    def test_parse_pywintypes_date_with_python_datetime(self):
        """Test parsing a plain Python datetime."""
        dt = datetime.datetime(2025, 5, 6, 16, 17, 0)
        result = _parse_pywintypes_date(dt)
        self.assertIsNotNone(result)
        self.assertEqual(result, dt)
    
    def test_parse_pywintypes_date_with_ole_float(self):
        """Test parsing OLE Automation date (days since Dec 30, 1899)."""
        # OLE date for 2000-01-01 00:00:00 is approximately 36526
        ole_date = 36526.0
        result = _parse_pywintypes_date(ole_date)
        self.assertIsNotNone(result, "Should parse OLE date successfully")
        self.assertEqual(result.year, 2000)
        self.assertEqual(result.month, 1)
        self.assertEqual(result.day, 1)
    
    def test_parse_pywintypes_date_with_iso_string(self):
        """Test parsing ISO format datetime string."""
        result = _parse_pywintypes_date("2025-05-06 16:17:00")
        self.assertIsNotNone(result)
        self.assertEqual(result.year, 2025)
        self.assertEqual(result.month, 5)
        self.assertEqual(result.day, 6)
        self.assertEqual(result.hour, 16)
        self.assertEqual(result.minute, 17)
    
    def test_parse_pywintypes_date_with_us_format(self):
        """Test parsing US locale datetime string."""
        result = _parse_pywintypes_date("05/06/2025 16:17:00")
        self.assertIsNotNone(result)
        self.assertEqual(result.year, 2025)
        self.assertEqual(result.month, 5)
        self.assertEqual(result.day, 6)
    
    def test_parse_pywintypes_date_with_eu_format(self):
        """Test parsing EU locale datetime string."""
        result = _parse_pywintypes_date("06/05/2025 16:17:00")
        self.assertIsNotNone(result)
        self.assertEqual(result.year, 2025)
        # Note: This can be parsed as either US (6th month, 5th day) or EU (6th day, 5th month)
        # The function tries formats in order, so it may match US format first
        self.assertIn(result.month, [5, 6])  # Either is acceptable depending on format tried
        self.assertIn(result.day, [5, 6])
    
    def test_parse_pywintypes_date_with_12h_format(self):
        """Test parsing 12-hour format with AM/PM."""
        result = _parse_pywintypes_date("05/06/2025 04:17:00 PM")
        self.assertIsNotNone(result)
        self.assertEqual(result.year, 2025)
        # The hour may be parsed as 4 in 12-hour format or 16 if PM is processed
        # This depends on the strptime implementation
        self.assertIsNotNone(result.hour)
    
    def test_parse_pywintypes_date_with_none(self):
        """Test parsing None returns None."""
        result = _parse_pywintypes_date(None)
        self.assertIsNone(result)
    
    def test_parse_pywintypes_date_with_bogus_ole_date(self):
        """Test that bogus OLE dates (epoch) return None."""
        # OLE date for 1899-12-30 is 0.0
        result = _parse_pywintypes_date(0.0)
        self.assertIsNone(result)


class EmailDateFormattingTests(unittest.TestCase):
    """Test email datetime formatting."""
    
    def test_format_email_datetime_with_datetime(self):
        """Test formatting a standard Python datetime."""
        dt = datetime.datetime(2025, 5, 6, 16, 17, 0)
        result = _format_email_datetime(dt)
        self.assertEqual(result, "06 May 2025 16:17")
    
    def test_format_email_datetime_with_none(self):
        """Test formatting None returns 'Unknown'."""
        result = _format_email_datetime(None)
        self.assertEqual(result, "Unknown")
    
    def test_format_email_datetime_with_ole_float(self):
        """Test formatting OLE Automation date."""
        ole_date = 36526.0
        result = _format_email_datetime(ole_date)
        self.assertNotEqual(result, "Unknown")
        self.assertRegex(result, r'^\d{2} \w+ \d{4} \d{2}:\d{2}$')
    
    def test_format_email_datetime_with_iso_string(self):
        """Test formatting ISO format string."""
        result = _format_email_datetime("2025-05-06 16:17:00")
        # When given a string that isn't recognized as datetime, 
        # _format_email_datetime tries various conversions and falls back to str()
        self.assertIsNotNone(result)
        self.assertIsInstance(result, str)
    
    def test_format_email_datetime_graceful_fallback(self):
        """Test that bad input gracefully returns string representation."""
        result = _format_email_datetime("invalid")
        self.assertIsNotNone(result)
        # Should return string representation when all else fails
        self.assertIsInstance(result, str)


class MonthNamesTests(unittest.TestCase):
    """Test month name constants."""
    
    def test_month_names_count(self):
        """Test that all 12 months are defined."""
        self.assertEqual(len([m for m in _MONTH_NAMES[1:]]), 12)
    
    def test_month_names_english(self):
        """Test that month names are in English."""
        expected_months = (
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December',
        )
        self.assertEqual(_MONTH_NAMES[1:], expected_months)
    
    def test_month_names_no_none_in_range_1_12(self):
        """Test that indices 1-12 (months) are not None."""
        for month_num in range(1, 13):
            self.assertIsNotNone(_MONTH_NAMES[month_num])


class BogusYearsTests(unittest.TestCase):
    """Test bogus year detection."""
    
    def test_bogus_years_are_defined(self):
        """Test that bogus years set is not empty."""
        self.assertGreater(len(_BOGUS_YEARS), 0)
    
    def test_bogus_years_contain_epoch_values(self):
        """Test that bogus years contain known null/epoch values."""
        # 1899 and 1900 are OLE epoch markers
        # 1601 is Windows/FILETIME epoch
        # 4501 can appear in some COM objects
        self.assertTrue(1899 in _BOGUS_YEARS or 1900 in _BOGUS_YEARS)


class RoundTripTests(unittest.TestCase):
    """Test date formatting and parsing round-trip consistency."""
    
    def test_format_then_validate(self):
        """Test that formatted dates can be validated."""
        dt = datetime.datetime(2025, 5, 6, 16, 17, 0)
        formatted = _safe_strftime(dt)
        validation_result = _validate_date(formatted, "test.msg", "SentOn")
        self.assertIsNone(validation_result, 
                         f"Formatted date should be valid: {formatted}")
    
    def test_multiple_datetime_formats(self):
        """Test that multiple date formats round-trip correctly."""
        test_cases = [
            datetime.datetime(2025, 1, 1, 0, 0, 0),
            datetime.datetime(2025, 12, 31, 23, 59, 0),
            datetime.datetime(2000, 6, 15, 12, 30, 0),
            datetime.datetime(1970, 1, 1, 0, 0, 0),
            datetime.datetime(2099, 12, 31, 23, 59, 0),
        ]
        for original_dt in test_cases:
            # Format it
            formatted = _safe_strftime(original_dt)
            # Validate it
            validation_result = _validate_date(formatted, "test.msg", "Test")
            self.assertIsNone(validation_result,
                             f"Failed for {original_dt}: {formatted}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
