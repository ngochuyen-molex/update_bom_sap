"""Value cleaning shared by all phases. No SAP, Excel or Snowflake code here."""
import pandas as pd


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def clean_component(value):
    value = clean_text(value)
    try:
        # handles (.0, .000)
        return str(int(float(value))).lstrip("0")
    except ValueError:
        return value.lstrip("0")

def clean_item(value):
    value = clean_text(value)

    if value.endswith(".0"):
        value = value[:-2]

    if value.isdigit():
        if len(value) == 1 and value != "0":
            value = str(int(value) * 10)
        if len(value) < 4:
            value = value.zfill(4)

    return value

def clean_number(value):
    if pd.isna(value):
        return None

    value = str(value).strip()

    if value == "":
        return None

    # SAP negative format
    # 160- -> -160
    if value.endswith("-"):
        value = "-" + value[:-1]

    # remove thousand separators
    value = value.replace(",", "")

    return float(value)

def numbers_equal(a, b, decimal_places=2):
    a = clean_number(a)
    b = clean_number(b)
    if a is None and b is None:
        return True

    if a is None or b is None:
        return False

    return round(a, decimal_places) == round(b, decimal_places)
