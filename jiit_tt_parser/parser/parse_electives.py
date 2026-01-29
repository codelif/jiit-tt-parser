import re
from typing import Dict

import pandas as pd


def parse_electives(
    file_path: str,
    sheet_name=0,
    header_row: int = 3,
    code_col_idx: int = 0,
    subj_col_idx: int = 3,
) -> Dict[str, str]:

    df = pd.read_excel(
        file_path, sheet_name=sheet_name, header=header_row, engine="openpyxl"
    ).dropna()

    code_col = df.columns[code_col_idx]
    subj_col = df.columns[subj_col_idx]

    pattern = re.compile(r"(.+?)(\d+)$")
    mapping = {}

    for raw_code, subject in zip(df[code_col].astype(str), df[subj_col].astype(str)):
        print(raw_code)
        raw_code = raw_code.strip()
        subject = subject.strip()

        if "/" in raw_code:
            base, _ = raw_code.split("/", 1)
            mapping[base] = subject

            m = pattern.match(base)
            if m:
                prefix, num = m.groups()
                next_code = f"{prefix}{int(num)+1}"
                mapping[next_code] = subject
        else:
            mapping[raw_code] = subject

    return mapping
