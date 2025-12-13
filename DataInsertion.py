import pandas as pd
import numpy as np

data_from_excel = pd.read_excel('')

def insert_data(df, data, index=None):
    """
    Inserts data into a DataFrame at the specified index.

    Parameters:
    df (pd.DataFrame): The original DataFrame.
    data (dict or pd.DataFrame): The data to insert.
    index (int, optional): The index at which to insert the data. 
                           If None, appends to the end.

    Returns:
    pd.DataFrame: The DataFrame with the new data inserted.
    """
    if isinstance(data, dict):
        new_data = pd.DataFrame([data])
    elif isinstance(data, pd.DataFrame):
        new_data = data
    else:
        raise ValueError("Data must be a dictionary or a pandas DataFrame.")

    if index is None:
        result_df = pd.concat([df, new_data], ignore_index=True)
    else:
        upper_half = df.iloc[:index]
        lower_half = df.iloc[index:]
        result_df = pd.concat([upper_half, new_data, lower_half], ignore_index=True)

    return result_df