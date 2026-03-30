import pandas as pd

def preprocess_input(df, categorical_features):
    """
    Ensures categorical features are strings for CatBoost
    """
    for col in categorical_features:
        df[col] = df[col].astype(str)
    return df
