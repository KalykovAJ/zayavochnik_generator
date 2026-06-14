import os
import pandas as pd
from typing import Union, List, Optional

def load_and_clean_data(
    file_path: str,
    filter_column: Optional[str],
    filter_value: Optional[Union[str, List[str]]],
    qty_col_name: str,
    total_col_name: str,
    weight_col_name: str
) -> pd.DataFrame:
    """Загружает исходник, фильтрует строки и готовит пустые целевые колонки."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Файл данных не найден по пути: {file_path}")

    df = pd.read_excel(file_path)

    if filter_column and filter_column in df.columns and filter_value is not None:
        exclude_list = [filter_value] if isinstance(filter_value, str) else filter_value
        df = df[~df[filter_column].isin(exclude_list)].copy()

    df[qty_col_name] = ""
    df[total_col_name] = ""
    df[weight_col_name] = ""

    return df