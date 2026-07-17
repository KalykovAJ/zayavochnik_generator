import os
import pandas as pd
from typing import Union, List, Optional

def load_and_clean_data(
    file_path: str,
    filter_column: Optional[Union[str, List[str]]],
    filter_value: Optional[Union[str, List[str]]],
    qty_col_name: str,
    total_col_name: str,
    weight_col_name: str
) -> pd.DataFrame:
    """Загружает исходник, фильтрует строки и готовит пустые целевые колонки."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Файл данных не найден по пути: {file_path}")

    df = pd.read_excel(file_path)

    if filter_column and filter_value is not None:
        columns_list = [filter_column] if isinstance(filter_column, str) else filter_column
        exclude_list = [filter_value] if isinstance(filter_value, str) else filter_value

        mask = pd.Series(True, index=df.index)
        for col in columns_list:
            if col not in df.columns:
                continue
            mask &= df[col].isin(exclude_list)
        df = df[mask].copy()

    df[qty_col_name] = ""
    df[total_col_name] = ""
    df[weight_col_name] = ""

    return df