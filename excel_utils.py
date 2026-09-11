import pandas as pd
import os

def load_excel(file_path):
    return pd.read_excel(file_path)


def save_excel(df, file_path):
    df.to_excel(file_path, index=False)


def prepare_result_columns(df):
    if "Status" not in df.columns:
        df["Status"] = ""

    if "Message" not in df.columns:
        df["Message"] = ""

    return df


def log_result(df, idx, status, message):
    df.loc[idx, "Status"] = status
    df.loc[idx, "Message"] = message



def add_log(df, idx, message):
    print(message)
    df.at[idx, "LOG"] += f"{message}\n"

def save_progress(input_file,df):
    temp_file = input_file.replace(".xlsx", "_temp.xlsx")
    df.to_excel(temp_file, index=False, engine="openpyxl")
    os.replace(temp_file, input_file)
