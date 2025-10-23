


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
import logging
import re
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Mapping, Tuple

import pandas as pd


LOGGER = logging.getLogger(__name__)

PERSIAN_DIGIT_MAP = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

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
    if not (1 <= j_month <= 12):
        raise ValueError(f"Invalid Jalali month: {j_month}")
    try:
        _ = JALALI_MONTH_DAYS[j_month - 1]
    except IndexError as exc:
        raise ValueError(f"Invalid Jalali month: {j_month}") from exc

    jy = j_year - 979
    jm = j_month - 1
    jd = j_day - 1

    if jm < 0 or jd < 0:
        raise ValueError("Invalid Jalali date components.")

    j_day_no = 365 * jy + jy // 33 * 8 + ((jy % 33) + 3) // 4
    for i in range(jm):
        j_day_no += JALALI_MONTH_DAYS[i]
    j_day_no += jd

    g_day_no = j_day_no + 79

    gy = 1600 + 400 * (g_day_no // 146097)
    g_day_no %= 146097

    leap = True
    if g_day_no >= 36525:
        g_day_no -= 1
        gy += 100 * (g_day_no // 36524)
        g_day_no %= 36524
        if g_day_no >= 365:
            g_day_no += 1
        else:
            leap = False

    gy += 4 * (g_day_no // 1461)
    g_day_no %= 1461

    if g_day_no >= 366:
        leap = False
        g_day_no -= 1
        gy += g_day_no // 365
        g_day_no %= 365

    gd = g_day_no + 1
    gregorian_month_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    gm = 0
    while gm < 12:
        days_in_month = gregorian_month_days[gm]
        if gm == 1 and leap:
            days_in_month += 1
        if gd <= days_in_month:
            break
        gd -= days_in_month
        gm += 1

    if gm >= 12:
        raise ValueError("Converted Gregorian month out of range.")

    return gy, gm + 1, gd


def parse_visit_date(value: object) -> pd.Timestamp:
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
    if re.fullmatch(r"\d{8}", text):
        text = f"{text[:4]}/{text[4:6]}/{text[6:]}"

    direct = pd.to_datetime(text, errors="coerce", dayfirst=False)
    if not pd.isna(direct) and direct.year >= 1900:
        return direct

    match = re.search(r"(\d{3,4})/(\d{1,2})/(\d{1,2})", text)
    if not match:
        return pd.NaT
    jy, jm, jd = map(int, match.groups())
    if jy < 1700:
        try:
            gy, gm, gd = jalali_to_gregorian(jy, jm, jd)
            return pd.Timestamp(datetime(gy, gm, gd))
        except ValueError:
            return pd.NaT
    return pd.NaT


def format_visit_date(value: pd.Timestamp) -> str | pd.NA:
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
    if len(digits) != 10:
        return pd.NA
    if len(set(digits)) == 1:
        return pd.NA
    checksum = int(digits[-1])
    total = sum(int(digits[i]) * (10 - i) for i in range(9))
    remainder = total % 11
    if (remainder < 2 and checksum == remainder) or (remainder >= 2 and checksum + remainder == 11):
        return digits
    return pd.NA


def clean_mobile(value: object) -> str | pd.NA:
    digits = re.sub(r"\D", "", normalize_digits(value))
    if digits.startswith("98") and len(digits) == 12:
        digits = "0" + digits[2:]
    if digits.startswith("0098") and len(digits) == 14:
        digits = "0" + digits[4:]
    if len(digits) == 10 and not digits.startswith("0"):
        digits = "0" + digits
    if len(digits) == 11 and digits.startswith("09"):
        return digits
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


def clean_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
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

    # Parse visit dates BEFORE filtering (needed for deduplication across all records)
    subset["visit_date_parsed"] = subset["visit_date_raw"].apply(parse_visit_date)

    # Sort by national_id and visit_date_parsed to prioritize most recent records
    subset = subset.sort_values(
        by=["national_id", "visit_date_parsed"], ascending=[True, False], na_position="last"
    )

    # Keep most recent record per national_id (based on تاریخ اخذ)
    before_dedup = len(subset)
    subset = subset.drop_duplicates(subset=["national_id"], keep="first").copy()
    if len(subset) != before_dedup:
        LOGGER.info("Kept most recent record per national_id, removed %d older duplicate records", before_dedup - len(subset))

    # Now apply name validation filters
    full_names = subset["full_name"].fillna("").astype(str)
    contains_karbar = full_names.str.contains("کاربر تلفنی", case=False, regex=False)
    contains_punctuation = full_names.str.contains(r"[.\-]", regex=True)
    contains_english = full_names.str.contains(r"[A-Za-z]", case=False, regex=True)

    # Check if first_name or last_name is less than 2 characters
    first_name_too_short = subset["first_name"].fillna("").astype(str).str.len() < 2
    last_name_too_short = subset["last_name"].fillna("").astype(str).str.len() < 2

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

    subset["visit_date"] = subset["visit_date_parsed"].apply(format_visit_date)
    subset["tags"] = subset.apply(build_tags, axis=1)

    cleaned_output = subset[["national_id", "first_name", "last_name", "mobile", "visit_date", "tags"]].copy()

    if not excluded.empty:
        excluded["visit_date_parsed"] = excluded["visit_date_raw"].apply(parse_visit_date)
        excluded["visit_date"] = excluded["visit_date_parsed"].apply(format_visit_date)
        excluded["tags"] = excluded.apply(build_tags, axis=1)
    else:
        excluded["visit_date"] = pd.NA
        excluded["tags"] = ""

    excluded_output = excluded[
        ["full_name", "national_id", "first_name", "last_name", "mobile", "visit_date", "tags"]
    ].copy()

    return cleaned_output, excluded_output


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
    cleaned, excluded = clean_dataframe(df)

    LOGGER.info("Writing %d rows to %s", len(cleaned), output)
    export_dataframe(cleaned, output)

    if not excluded.empty:
        excluded_output = output.with_name(f"{output.stem}_excluded{output.suffix}")
        LOGGER.info("Writing %d excluded rows to %s", len(excluded), excluded_output)
        export_dataframe(excluded, excluded_output)


if __name__ == "__main__":
    main()
