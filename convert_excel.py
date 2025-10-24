


"""
Utility script to reshape Noor Queue export files into the format expected by
the Noor appointment importer.

Input Excel columns (Farsi):
    اپراتور، درمانگاه، پزشک، تاریخ اخذ، ساعت اخذ، تاریخ ویزیت، ساعت، پذیرش،
    نوبت، بیمار، پیگیری، کدملی، موبایل، CallerID، نوع، پرداخت، مبلغ، کد بانک،
    توضیحات، وضعیت

Output columns (English):
    national_id (required), first_name (required), last_name (required),
    mobile (optional), tags (autofilled)
"""

from __future__ import annotations

import argparse
import json
import logging
import random
import re
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Mapping, Tuple

import pandas as pd
import jdatetime


LOGGER = logging.getLogger(__name__)

PERSIAN_DIGIT_MAP = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

# Load Persian names gender database
def _load_persian_names_gender() -> dict:
    """Load Persian names gender database from CSV and JSON files."""
    gender_lookup = {}

    # First, try to load from CSV file (higher priority)
    try:
        csv_path = Path("iranian_names_full.csv")
        if csv_path.exists():
            df = pd.read_csv(csv_path)
            for _, row in df.iterrows():
                name_fa = str(row['name_fa']).strip()
                gender = str(row['gender']).strip().lower()
                if name_fa and gender in ['male', 'female']:
                    gender_lookup[name_fa] = gender

            LOGGER.info("Loaded %d Persian names from CSV file", len(gender_lookup))
        else:
            LOGGER.warning("Iranian names CSV file not found: %s", csv_path)
    except Exception as e:
        LOGGER.error("Error loading Persian names from CSV: %s", e)

    # Then, try to load from JSON file (lower priority, only if not already in lookup)
    try:
        json_path = Path("persian_names_gender.json")
        if json_path.exists():
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Add male names (only if not already in lookup)
            for name in data.get("male", []):
                if name not in gender_lookup:
                    gender_lookup[name] = "male"

            # Add female names (only if not already in lookup)
            for name in data.get("female", []):
                if name not in gender_lookup:
                    gender_lookup[name] = "female"

            # Add unisex names (randomly assign male/female, only if not already in lookup)
            for name in data.get("unisex", []):
                if name not in gender_lookup:
                    gender_lookup[name] = random.choice(["male", "female"])

            LOGGER.info("Added %d additional names from JSON file",
                       len([name for name in data.get("male", []) + data.get("female", []) + data.get("unisex", [])
                            if name not in gender_lookup]))
        else:
            LOGGER.warning("Persian names JSON file not found: %s", json_path)
    except Exception as e:
        LOGGER.error("Error loading Persian names from JSON: %s", e)

    if not gender_lookup:
        LOGGER.warning("No Persian names gender database found. Gender detection will be disabled.")

    return gender_lookup

# Global gender lookup dictionary
PERSIAN_GENDER_LOOKUP = _load_persian_names_gender()

COLUMN_ALIASES: Mapping[str, Tuple[str, ...]] = {
    "national_id": ("کدملی", "کد ملی", "شناسه ملی"),
    "full_name": ("بیمار", "نام و نام خانوادگی", "نام بیمار"),
    "mobile": ("موبایل", "شماره موبایل", "شماره تماس"),
}

OPTIONAL_COLUMN_ALIASES: Mapping[str, Tuple[str, ...]] = {
    "visit_date": ("تاریخ اخذ",),
    "status": ("وضعیت",),
    "appointment_type": ("نوع",),
    "clinic": ("درمانگاه",),
}

LETTER_NORMALIZATION_MAP = str.maketrans({"ي": "ی", "ك": "ک"})

STATUS_TAG_SOURCE: Mapping[str, str] = {
    "ثبت نوبت": "not_showed_patient",
    "چاپ نوبت": "showup_patient",
    "کنسل شده": "canceling_patient",
}

APPOINTMENT_TYPE_TAG_SOURCE: Mapping[str, str] = {
    "اينترنتي": "internet_user",
    "اینترنتی": "internet_user",
    "فالوآپ": "phone_user",
    "فالو آپ": "phone_user",
}

CLINIC_TAG_SOURCE: Mapping[str, str] = {
    "کلینیک  ویژه فوق تخصصی جراحی چاقی": "bariatric_surgery_clinic",
    "کلینیک  ویژه فوق تخصصی چکاپ": "checkup_specialty_clinic",
    "کلینیک  ویژه فوق تخصصی گوارش و کبد": "gastro_hepatology_specialty_clinic",
    "کلینیک بینایی سنجی": "optometry_clinic",
    "کلینیک تخصصی ارتوپدی": "orthopedics_clinic",
    "کلینیک تخصصی اورولوژی(جراحی کلیه و پروستات)": "urology_clinic",
    "کلینیک تخصصی بیماری های داخلی": "internal_medicine_clinic",
    "کلینیک تخصصی جراحی اطفال": "pediatric_surgery_clinic",
    "کلینیک تخصصی جراحی جنرال": "general_surgery_clinic",
    "کلینیک تخصصی جراحی زنان و زایمان": "obgyn_surgery_clinic",
    "کلینیک تخصصی جراحی عروق و واریس": "vascular_varicose_surgery_clinic",
    "کلینیک تخصصی جراحی فک و صورت": "oral_maxillofacial_surgery_clinic",
    "کلینیک تخصصی جراحی قلب": "cardiac_surgery_clinic",
    "کلینیک تخصصی جراحی مغز و اعصاب": "neurosurgery_clinic",
    "کلینیک تخصصی خون و سرطان": "hematology_oncology_clinic",
    "کلینیک تخصصی داخلی": "internal_specialty_clinic",
    "کلینیک تخصصی داخلی اطفال و نوزادان": "pediatric_neonatal_internal_clinic",
    "کلینیک تخصصی داخلی ریه": "pulmonology_clinic",
    "کلینیک تخصصی داخلی قلب و عروق": "cardiology_clinic",
    "کلینیک تخصصی داخلی مغز و اعصاب اطفال": "pediatric_neurology_clinic",
    "کلینیک تخصصی روانپزشکی": "psychiatry_clinic",
    "کلینیک تخصصی زیبایی و بیوتی": "aesthetics_beauty_clinic",
    "کلینیک تخصصی عفونی": "infectious_diseases_clinic",
    "کلینیک تخصصی پوست، مو و زیبایی": "dermatology_hair_aesthetics_clinic",
    "کلینیک تخصصی چشم": "ophthalmology_clinic",
    "کلینیک تخصصی گوش و حلق و بینی": "ent_clinic",
    "کلینیک تغذیه و رژیم درمانی": "nutrition_diet_therapy_clinic",
    "کلینیک جراحی پلاستیک": "plastic_surgery_clinic",
    "کلینیک داخلی مغز و اعصاب": "neurology_clinic",
    "کلینیک زخم": "wound_care_clinic",
    "کلینیک شنوایی سنجی": "audiology_clinic",
    "کلینیک فوق تخصصی اختلالات جنسی و زناشویی": "sexual_marital_disorders_clinic",
    "کلینیک فوق تخصصی بیماری های قلب": "cardiac_super_specialty_clinic",
    "کلینیک فوق تخصصی نفرولوژی، غدد، دیابت و تیروئید": "nephrology_endocrine_diabetes_thyroid_clinic",
    "کلینیک فوق تخصصی پستان": "breast_super_specialty_clinic",
    "کلینیک فوق تخصصی گوارش اطفال": "pediatric_gastroenterology_clinic",
    "کلینیک مشاوره و روانشناسی": "counseling_psychology_clinic",
    "کلینیک ویژه  فوق تخصصی درد": "pain_super_specialty_clinic",
    "کلینیک ویژه  فوق تخصصی زانو و تعویض مفصل": "knee_joint_replacement_clinic",
    "کلینیک ویژه  فوق تخصصی قلب اطفال تا 15 سال": "pediatric_cardiology_clinic",
    "کلینیک ویژه فوق تخصصی روماتولوژی": "rheumatology_super_specialty_clinic",
    "کلینیک ویژه فوق تخصصی قلب": "heart_super_specialty_clinic",
}

JALALI_MONTH_DAYS = (31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)


def _normalize_header(value: str) -> str:
    return re.sub(r"\s+", "", str(value)).casefold()


def _select_columns(columns: Iterable[str]) -> Mapping[str, str]:
    normalized = {_normalize_header(col): col for col in columns}
    selected = {}
    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            candidate = normalized.get(_normalize_header(alias))
            if candidate:
                selected[target] = candidate
                break
        if target not in selected:
            raise KeyError(f"Cannot find column for {target!r}; make sure one of {aliases} exists.")
    return selected


def _select_optional_columns(columns: Iterable[str]) -> Mapping[str, str]:
    normalized = {_normalize_header(col): col for col in columns}
    selected: dict[str, str] = {}
    for target, aliases in OPTIONAL_COLUMN_ALIASES.items():
        for alias in aliases:
            candidate = normalized.get(_normalize_header(alias))
            if candidate:
                selected[target] = candidate
                break
    return selected


def normalize_digits(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if not text or text.casefold() in {"nan", "none"}:
        return ""
    return text.translate(PERSIAN_DIGIT_MAP)


def normalize_farsi_text(value: object) -> str:
    text = normalize_digits(value)
    if not text:
        return ""
    normalized = text.translate(LETTER_NORMALIZATION_MAP)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def _prepare_normalized_mapping(mapping: Mapping[str, str]) -> dict[str, str]:
    return {normalize_farsi_text(key).casefold(): value for key, value in mapping.items()}


STATUS_TAG_MAP = _prepare_normalized_mapping(STATUS_TAG_SOURCE)
APPOINTMENT_TYPE_TAG_MAP = _prepare_normalized_mapping(APPOINTMENT_TYPE_TAG_SOURCE)
CLINIC_TAG_MAP = _prepare_normalized_mapping(CLINIC_TAG_SOURCE)


def jalali_to_gregorian(j_year: int, j_month: int, j_day: int) -> Tuple[int, int, int]:
    """
    Convert Jalali date to Gregorian using jdatetime library.
    """
    try:
        jalali_date = jdatetime.date(j_year, j_month, j_day)
        gregorian_date = jalali_date.togregorian()
        return gregorian_date.year, gregorian_date.month, gregorian_date.day
    except (ValueError, TypeError) as e:
        raise ValueError(f"Invalid Jalali date: {j_year}/{j_month}/{j_day}") from e


def parse_visit_date(value: object) -> pd.Timestamp:
    """
    Enhanced date parsing that:
    1. Detects Jalali vs Gregorian dates automatically
    2. Converts Jalali to Gregorian for database storage
    3. Handles time components (defaults to 00:00 if missing)
    4. Validates date ranges appropriately
    """
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value.tz_localize(None) if value.tzinfo else value
    if isinstance(value, (datetime, date)):
        return pd.to_datetime(value, errors="coerce")
    if isinstance(value, (int, float)) and not pd.isna(value):
        excel_candidate = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if not pd.isna(excel_candidate):
            return excel_candidate

    text = normalize_digits(value)
    text = re.sub(r"[\u200c\u200f]", "", text)
    text = text.replace(".", "/").replace("-", "/").strip()
    if not text:
        return pd.NaT

    # Handle 8-digit format (YYYYMMDD)
    if re.fullmatch(r"\d{8}", text):
        text = f"{text[:4]}/{text[4:6]}/{text[6:]}"

    # Try to parse as datetime with time component first
    datetime_patterns = [
        r"(\d{3,4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{1,2})",  # YYYY/MM/DD HH:MM
        r"(\d{3,4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})",  # YYYY/MM/DD HH:MM:SS
    ]

    for pattern in datetime_patterns:
        match = re.search(pattern, text)
        if match:
            groups = match.groups()
            year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
            hour = int(groups[3]) if len(groups) > 3 else 0
            minute = int(groups[4]) if len(groups) > 4 else 0
            second = int(groups[5]) if len(groups) > 5 else 0

            # Determine if it's Jalali or Gregorian
            if _is_jalali_date(year, month, day):
                try:
                    gy, gm, gd = jalali_to_gregorian(year, month, day)
                    return pd.Timestamp(datetime(gy, gm, gd, hour, minute, second))
                except ValueError:
                    return pd.NaT
            else:
                # Gregorian date
                try:
                    return pd.Timestamp(datetime(year, month, day, hour, minute, second))
                except ValueError:
                    return pd.NaT

    # Try to parse as date only (no time)
    match = re.search(r"(\d{3,4})/(\d{1,2})/(\d{1,2})", text)
    if not match:
        return pd.NaT

    year, month, day = map(int, match.groups())

    # Determine if it's Jalali or Gregorian
    if _is_jalali_date(year, month, day):
        try:
            gy, gm, gd = jalali_to_gregorian(year, month, day)
            # Default to 00:00:00 if no time specified
            return pd.Timestamp(datetime(gy, gm, gd, 0, 0, 0))
        except ValueError:
            return pd.NaT
    else:
        # Gregorian date - default to 00:00:00 if no time specified
        try:
            return pd.Timestamp(datetime(year, month, day, 0, 0, 0))
        except ValueError:
            return pd.NaT


def _is_jalali_date(year: int, month: int, day: int) -> bool:
    """
    Determine if a date is likely Jalali using jdatetime library validation.
    """
    # Jalali years are typically in the 1300-1500 range
    if year < 1300 or year > 1500:
        return False

    # Use jdatetime to validate if it's a valid Jalali date
    try:
        jdatetime.date(year, month, day)
        return True
    except (ValueError, TypeError):
        return False


def format_visit_date(value: pd.Timestamp) -> str | pd.NA:
    """
    Format date for database storage (always Gregorian/ISO format).
    This ensures data integrity by storing dates in a standard format.
    """
    if pd.isna(value):
        return pd.NA
    if isinstance(value, pd.Timestamp):
        ts = value.tz_localize(None) if value.tzinfo else value
    elif isinstance(value, (datetime, date)):
        ts = pd.Timestamp(value)
    else:
        ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return pd.NA
    return ts.strftime("%Y-%m-%d")


def format_visit_date_for_ui(value: pd.Timestamp) -> str | pd.NA:
    """
    Format date for UI display (Jalali format for better user experience).
    This converts Gregorian dates back to Jalali for display purposes.
    """
    if pd.isna(value):
        return pd.NA
    if isinstance(value, pd.Timestamp):
        ts = value.tz_localize(None) if value.tzinfo else value
    elif isinstance(value, (datetime, date)):
        ts = pd.Timestamp(value)
    else:
        ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return pd.NA

    # Convert Gregorian to Jalali for display
    try:
        jy, jm, jd = gregorian_to_jalali(ts.year, ts.month, ts.day)
        return f"{jy:04d}/{jm:02d}/{jd:02d}"
    except (ValueError, TypeError):
        # Fallback to Gregorian if conversion fails
        return ts.strftime("%Y-%m-%d")


def format_visit_datetime_for_ui(value: pd.Timestamp) -> str | pd.NA:
    """
    Format datetime for UI display (Jalali format with time).
    This converts Gregorian datetime back to Jalali for display purposes.
    """
    if pd.isna(value):
        return pd.NA
    if isinstance(value, pd.Timestamp):
        ts = value.tz_localize(None) if value.tzinfo else value
    elif isinstance(value, (datetime, date)):
        ts = pd.Timestamp(value)
    else:
        ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return pd.NA

    # Convert Gregorian to Jalali for display
    try:
        jy, jm, jd = gregorian_to_jalali(ts.year, ts.month, ts.day)
        return f"{jy:04d}/{jm:02d}/{jd:02d} {ts.hour:02d}:{ts.minute:02d}"
    except (ValueError, TypeError):
        # Fallback to Gregorian if conversion fails
        return ts.strftime("%Y-%m-%d %H:%M")


def format_visit_date_for_database(value: pd.Timestamp) -> str | pd.NA:
    """
    Format date for database storage (ISO format with timezone).
    This ensures proper database storage and querying capabilities.
    """
    if pd.isna(value):
        return pd.NA
    if isinstance(value, pd.Timestamp):
        ts = value.tz_localize(None) if value.tzinfo else value
    elif isinstance(value, (datetime, date)):
        ts = pd.Timestamp(value)
    else:
        ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return pd.NA

    # Return ISO format for database storage
    return ts.isoformat()


def gregorian_to_jalali(g_year: int, g_month: int, g_day: int) -> Tuple[int, int, int]:
    """
    Convert Gregorian date to Jalali using jdatetime library.
    """
    try:
        gregorian_date = datetime(g_year, g_month, g_day).date()
        jalali_date = jdatetime.date.fromgregorian(date=gregorian_date)
        return jalali_date.year, jalali_date.month, jalali_date.day
    except (ValueError, TypeError) as e:
        raise ValueError(f"Invalid Gregorian date: {g_year}/{g_month}/{g_day}") from e


def build_tags(row: pd.Series) -> str:
    tags: list[str] = []

    def add(tag: str | None) -> None:
        if tag and tag not in tags:
            tags.append(tag)

    add("noor_hospital_queue")
    add("patient")

    status_value = normalize_farsi_text(row.get("status_raw", pd.NA))
    if status_value:
        add(STATUS_TAG_MAP.get(status_value.casefold()))

    appointment_type_value = normalize_farsi_text(row.get("appointment_type_raw", pd.NA))
    if appointment_type_value:
        add(APPOINTMENT_TYPE_TAG_MAP.get(appointment_type_value.casefold()))

    clinic_value = normalize_farsi_text(row.get("clinic_raw", pd.NA))
    if clinic_value:
        add(CLINIC_TAG_MAP.get(clinic_value.casefold()))

    if not tags:
        return ""
    return ",".join(tags) + ","


def clean_national_id(value: object) -> str | pd.NA:
    digits = re.sub(r"\D", "", normalize_digits(value))
    if len(digits) < 8 or len(digits) > 11:
        return pd.NA
    if len(set(digits)) == 1:
        return pd.NA

    # More lenient validation - accept 8-11 digit national IDs
    # This allows more records to pass through while still filtering obvious invalid IDs
    try:
        # Basic validation: should be 8-11 digits, not all the same
        if 8 <= len(digits) <= 11 and len(set(digits)) > 1:
            return digits
    except (ValueError, TypeError):
        pass
    return pd.NA


def clean_mobile(value: object) -> str | pd.NA:
    digits = re.sub(r"\D", "", normalize_digits(value))
    if not digits:
        return pd.NA

    # More lenient mobile validation
    # Handle different formats
    if digits.startswith("98") and len(digits) == 12:
        digits = "0" + digits[2:]
    if digits.startswith("0098") and len(digits) == 14:
        digits = "0" + digits[4:]
    if len(digits) == 10 and not digits.startswith("0"):
        digits = "0" + digits

    # Accept various valid Iranian mobile formats
    if len(digits) == 11 and digits.startswith("09"):
        return digits
    elif len(digits) == 10 and digits.startswith("9"):
        return "0" + digits
    elif len(digits) == 10:
        return digits  # Accept 10-digit numbers as valid
    elif len(digits) == 11:
        return digits  # Accept 11-digit numbers as valid

    return pd.NA


def split_full_name(value: object) -> Tuple[str | pd.NA, str | pd.NA]:
    if value is None or pd.isna(value):
        return pd.NA, pd.NA
    text = str(value).strip()
    if not text or text.casefold() in {"nan", "none"}:
        return pd.NA, pd.NA
    pieces = [part for part in re.split(r"\s+", text) if part]
    if not pieces:
        return pd.NA, pd.NA
    if len(pieces) == 1:
        return pieces[0], pd.NA
    first = " ".join(pieces[:-1])
    last = pieces[-1]
    return first, last


def detect_gender(first_name: object) -> str | pd.NA:
    """
    Detect gender based on Persian first name using local database lookup.
    Returns 'male', 'female', or pd.NA if not found.
    """
    if first_name is None or pd.isna(first_name):
        return pd.NA

    # Normalize the name
    name = normalize_farsi_text(first_name).strip()
    if not name:
        return pd.NA

    # Look up in the gender database
    gender = PERSIAN_GENDER_LOOKUP.get(name)
    if gender:
        return gender

    # Try with first word only (in case of compound names)
    first_word = name.split()[0] if name.split() else name
    gender = PERSIAN_GENDER_LOOKUP.get(first_word)
    if gender:
        return gender

    # Try with common Persian name patterns
    # Remove common prefixes/suffixes and try again
    cleaned_name = re.sub(r'^(آقای|خانم|دکتر|مهندس|استاد|جناب|سرکار|سرکار خانم|آقا|خانم)\s*', '', name)
    if cleaned_name != name:
        gender = PERSIAN_GENDER_LOOKUP.get(cleaned_name)
        if gender:
            return gender

    return pd.NA


def _is_name_complete(first_name: str, last_name: str) -> bool:
    """Check if name is complete (both first and last name have at least 3 characters)."""
    if pd.isna(first_name) or pd.isna(last_name):
        return False
    first_str = str(first_name).strip()
    last_str = str(last_name).strip()
    return len(first_str) >= 3 and len(last_str) >= 3


def _enhanced_deduplication(df: pd.DataFrame) -> pd.DataFrame:
    """
    Enhanced deduplication that:
    1. Keeps most recent record per national_id
    2. If most recent record has incomplete name, looks for earlier records with complete name
    3. Returns records with incomplete names that couldn't be completed
    """
    result_records = []
    incomplete_records = []

    # Group by national_id
    for national_id, group in df.groupby("national_id"):
        if pd.isna(national_id):
            continue

        # Sort group by visit_date_parsed (most recent first)
        group = group.sort_values("visit_date_parsed", ascending=False, na_position="last")

        # Get the most recent record
        most_recent = group.iloc[0]

        # Check if most recent record has complete name
        if _is_name_complete(most_recent["first_name"], most_recent["last_name"]):
            # Use most recent record
            result_records.append(most_recent)
        else:
            # Look for earlier records with complete name
            found_complete = False
            for _, record in group.iterrows():
                if _is_name_complete(record["first_name"], record["last_name"]):
                    # Use the most recent record but with complete name from earlier record
                    updated_record = most_recent.copy()
                    updated_record["first_name"] = record["first_name"]
                    updated_record["last_name"] = record["last_name"]
                    updated_record["full_name"] = f"{record['first_name']} {record['last_name']}"
                    result_records.append(updated_record)
                    found_complete = True
                    break

            if not found_complete:
                # No complete name found, add to incomplete records
                incomplete_records.append(most_recent)

    # Convert to DataFrames
    if result_records:
        result_df = pd.DataFrame(result_records)
    else:
        result_df = pd.DataFrame(columns=df.columns)

    if incomplete_records:
        incomplete_df = pd.DataFrame(incomplete_records)
        # Store incomplete records globally for later use
        _enhanced_deduplication.incomplete_records = incomplete_df
    else:
        _enhanced_deduplication.incomplete_records = pd.DataFrame(columns=df.columns)

    LOGGER.info("Enhanced deduplication: %d complete records, %d incomplete records",
                len(result_df), len(incomplete_records))

    return result_df


def clean_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Enhanced cleaning function that returns 4 dataframes:
    1. cleaned_output: Valid records with complete names
    2. excluded_output: Records with invalid/incomplete names
    3. duplicate_phone_output: Records with duplicate phone numbers
    4. incomplete_name_output: Records where name couldn't be completed from earlier records
    """
    selectors = _select_columns(df.columns)
    optional_selectors = _select_optional_columns(df.columns)

    column_order = [
        selectors["national_id"],
        selectors["full_name"],
        selectors["mobile"],
    ]
    for key in ("visit_date", "status", "appointment_type", "clinic"):
        if key in optional_selectors:
            column_order.append(optional_selectors[key])
    # Remove duplicates while preserving order
    column_order = list(dict.fromkeys(column_order))

    subset = df[column_order].copy()

    rename_map = {
        selectors["national_id"]: "national_id_raw",
        selectors["full_name"]: "full_name",
        selectors["mobile"]: "mobile_raw",
    }
    if "visit_date" in optional_selectors:
        rename_map[optional_selectors["visit_date"]] = "visit_date_raw"
    if "status" in optional_selectors:
        rename_map[optional_selectors["status"]] = "status_raw"
    if "appointment_type" in optional_selectors:
        rename_map[optional_selectors["appointment_type"]] = "appointment_type_raw"
    if "clinic" in optional_selectors:
        rename_map[optional_selectors["clinic"]] = "clinic_raw"
    subset = subset.rename(columns=rename_map)

    for column in ("visit_date_raw", "status_raw", "appointment_type_raw", "clinic_raw"):
        if column not in subset.columns:
            subset[column] = pd.NA

    subset["national_id"] = subset["national_id_raw"].apply(clean_national_id)
    subset["mobile"] = subset["mobile_raw"].apply(clean_mobile)

    name_parts = subset["full_name"].apply(
        lambda val: pd.Series(split_full_name(val), index=["first_name", "last_name"])
    )
    subset = subset.join(name_parts)

    # Add gender detection
    subset["gender"] = subset["first_name"].apply(detect_gender)

    # Parse visit dates BEFORE filtering (needed for deduplication across all records)
    subset["visit_date_parsed"] = subset["visit_date_raw"].apply(parse_visit_date)

    # Sort by national_id and visit_date_parsed to prioritize most recent records
    subset = subset.sort_values(
        by=["national_id", "visit_date_parsed"], ascending=[True, False], na_position="last"
    )

    # Enhanced deduplication logic with name completion
    subset = _enhanced_deduplication(subset)

    # Now apply name validation filters
    full_names = subset["full_name"].fillna("").astype(str)
    contains_karbar = full_names.str.contains("کاربر تلفنی", case=False, regex=False)
    contains_punctuation = full_names.str.contains(r"[.\-]", regex=True)
    contains_english = full_names.str.contains(r"[A-Za-z]", case=False, regex=True)

    # Check if first_name or last_name is less than 3 characters (enhanced requirement)
    first_name_too_short = subset["first_name"].fillna("").astype(str).str.len() < 3
    last_name_too_short = subset["last_name"].fillna("").astype(str).str.len() < 3

    # Check if first_name or last_name contains only digits
    first_name_numeric = subset["first_name"].fillna("").astype(str).str.match(r"^\d+$")
    last_name_numeric = subset["last_name"].fillna("").astype(str).str.match(r"^\d+$")

    invalid_name_mask = (
        contains_karbar | contains_punctuation | contains_english |
        first_name_too_short | last_name_too_short |
        first_name_numeric | last_name_numeric
    )

    excluded = subset[invalid_name_mask].copy()
    if not excluded.empty:
        LOGGER.info("Moved %d rows with invalid names to excluded set", len(excluded))

    subset = subset[~invalid_name_mask].copy()

    required_mask = subset["national_id"].notna() & subset["first_name"].notna() & subset["last_name"].notna()
    dropped = len(subset) - int(required_mask.sum())
    if dropped:
        LOGGER.warning("Dropping %d rows missing required fields", dropped)
    subset = subset[required_mask].copy()

    # Check for duplicate phone numbers
    phone_duplicates = subset[subset["mobile"].notna() & subset.duplicated(subset=["mobile"], keep=False)]
    if not phone_duplicates.empty:
        LOGGER.info("Found %d records with duplicate phone numbers", len(phone_duplicates))
        # Remove duplicates from main dataset
        subset = subset.drop(phone_duplicates.index)

    # Format dates for different purposes
    subset["visit_date"] = subset["visit_date_parsed"].apply(format_visit_date)  # Database storage (Gregorian)
    subset["visit_date_ui"] = subset["visit_date_parsed"].apply(format_visit_date_for_ui)  # UI display (Jalali)
    subset["visit_datetime_ui"] = subset["visit_date_parsed"].apply(format_visit_datetime_for_ui)  # UI display with time
    subset["visit_date_db"] = subset["visit_date_parsed"].apply(format_visit_date_for_database)  # Database ISO format
    subset["tags"] = subset.apply(build_tags, axis=1)

    cleaned_output = subset[["national_id", "first_name", "last_name", "gender", "mobile", "visit_date", "visit_date_ui", "visit_datetime_ui", "visit_date_db", "tags"]].copy()

    # Process excluded records
    if not excluded.empty:
        excluded["visit_date_parsed"] = excluded["visit_date_raw"].apply(parse_visit_date)
        excluded["visit_date"] = excluded["visit_date_parsed"].apply(format_visit_date)
        excluded["visit_date_ui"] = excluded["visit_date_parsed"].apply(format_visit_date_for_ui)
        excluded["visit_datetime_ui"] = excluded["visit_date_parsed"].apply(format_visit_datetime_for_ui)
        excluded["visit_date_db"] = excluded["visit_date_parsed"].apply(format_visit_date_for_database)
        excluded["tags"] = excluded.apply(build_tags, axis=1)
    else:
        excluded["visit_date"] = pd.NA
        excluded["visit_date_ui"] = pd.NA
        excluded["visit_datetime_ui"] = pd.NA
        excluded["visit_date_db"] = pd.NA
        excluded["tags"] = ""

    excluded_output = excluded[
        ["full_name", "national_id", "first_name", "last_name", "gender", "mobile", "visit_date", "visit_date_ui", "visit_datetime_ui", "visit_date_db", "tags"]
    ].copy()

    # Process duplicate phone records
    if not phone_duplicates.empty:
        phone_duplicates["visit_date"] = phone_duplicates["visit_date_parsed"].apply(format_visit_date)
        phone_duplicates["visit_date_ui"] = phone_duplicates["visit_date_parsed"].apply(format_visit_date_for_ui)
        phone_duplicates["visit_datetime_ui"] = phone_duplicates["visit_date_parsed"].apply(format_visit_datetime_for_ui)
        phone_duplicates["visit_date_db"] = phone_duplicates["visit_date_parsed"].apply(format_visit_date_for_database)
        phone_duplicates["tags"] = phone_duplicates.apply(build_tags, axis=1)
    else:
        phone_duplicates = pd.DataFrame(columns=["national_id", "first_name", "last_name", "gender", "mobile", "visit_date", "visit_date_ui", "visit_datetime_ui", "visit_date_db", "tags"])

    duplicate_phone_output = phone_duplicates[
        ["national_id", "first_name", "last_name", "gender", "mobile", "visit_date", "visit_date_ui", "visit_datetime_ui", "visit_date_db", "tags"]
    ].copy()

    # Process incomplete name records from enhanced deduplication
    if hasattr(_enhanced_deduplication, 'incomplete_records') and not _enhanced_deduplication.incomplete_records.empty:
        incomplete_records = _enhanced_deduplication.incomplete_records.copy()
        incomplete_records["visit_date"] = incomplete_records["visit_date_parsed"].apply(format_visit_date)
        incomplete_records["visit_date_ui"] = incomplete_records["visit_date_parsed"].apply(format_visit_date_for_ui)
        incomplete_records["visit_datetime_ui"] = incomplete_records["visit_date_parsed"].apply(format_visit_datetime_for_ui)
        incomplete_records["visit_date_db"] = incomplete_records["visit_date_parsed"].apply(format_visit_date_for_database)
        incomplete_records["tags"] = incomplete_records.apply(build_tags, axis=1)
        incomplete_name_output = incomplete_records[
            ["national_id", "first_name", "last_name", "gender", "mobile", "visit_date", "visit_date_ui", "visit_datetime_ui", "visit_date_db", "tags"]
        ].copy()
        LOGGER.info("Found %d records with incomplete names that couldn't be completed", len(incomplete_name_output))
    else:
        incomplete_name_output = pd.DataFrame(columns=["national_id", "first_name", "last_name", "gender", "mobile", "visit_date", "visit_date_ui", "visit_datetime_ui", "visit_date_db", "tags"])

    return cleaned_output, excluded_output, duplicate_phone_output, incomplete_name_output


def merge_dataframes(input_files: list[Path]) -> pd.DataFrame:
    """Read and merge multiple Excel files into a single DataFrame."""
    all_dataframes: list[pd.DataFrame] = []

    for input_file in input_files:
        if not input_file.exists():
            raise FileNotFoundError(f"Cannot find input file: {input_file}")

        LOGGER.info("Reading input file %s", input_file)
        df = pd.read_excel(input_file)
        all_dataframes.append(df)

    if not all_dataframes:
        raise ValueError("No input files provided")

    if len(all_dataframes) == 1:
        return all_dataframes[0]

    LOGGER.info("Merging %d input files", len(all_dataframes))
    merged_df = pd.concat(all_dataframes, ignore_index=True)
    LOGGER.info("Total rows after merge: %d", len(merged_df))

    return merged_df


def export_dataframe(df: pd.DataFrame, output_path: Path) -> None:
    if output_path.suffix.casefold() in {".xlsx", ".xlsm", ".xls"}:
        df.to_excel(output_path, index=False)
    elif output_path.suffix.casefold() == ".csv":
        df.to_csv(output_path, index=False)
    else:
        raise ValueError(
            f"Unsupported output format '{output_path.suffix}'. Use '.xlsx', '.xls', '.xlsm', or '.csv'."
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Noor queue exports into a simplified template.",
    )
    parser.add_argument("input", type=Path, nargs='+', help="Path(s) to the source Excel file(s). Multiple files will be merged.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Destination file (defaults to <input_stem>_cleaned.xlsx).",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow overwriting an existing output file.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Set the verbosity of log messages (default: INFO).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level.upper()), format="%(levelname)s: %(message)s")

    # Merge all input files
    df = merge_dataframes(args.input)

    output = args.output
    if output is None:
        if len(args.input) == 1:
            output = args.input[0].with_name(f"{args.input[0].stem}_cleaned.xlsx")
        else:
            output = Path("merged_cleaned.xlsx")

    if output.exists() and not args.overwrite:
        raise FileExistsError(f"Output file already exists: {output}. Use --overwrite to replace it.")

    LOGGER.info("Cleaning data")
    cleaned, excluded, duplicate_phone, incomplete_name = clean_dataframe(df)

    LOGGER.info("Writing %d rows to %s", len(cleaned), output)
    export_dataframe(cleaned, output)

    if not excluded.empty:
        excluded_output = output.with_name(f"{output.stem}_excluded{output.suffix}")
        LOGGER.info("Writing %d excluded rows to %s", len(excluded), excluded_output)
        export_dataframe(excluded, excluded_output)

    if not duplicate_phone.empty:
        duplicate_phone_output = output.with_name(f"{output.stem}_duplicate_phone{output.suffix}")
        LOGGER.info("Writing %d duplicate phone records to %s", len(duplicate_phone), duplicate_phone_output)
        export_dataframe(duplicate_phone, duplicate_phone_output)

    if not incomplete_name.empty:
        incomplete_name_output = output.with_name(f"{output.stem}_incomplete_name{output.suffix}")
        LOGGER.info("Writing %d incomplete name records to %s", len(incomplete_name), incomplete_name_output)
        export_dataframe(incomplete_name, incomplete_name_output)


if __name__ == "__main__":
    main()
