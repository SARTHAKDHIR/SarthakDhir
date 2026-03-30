import pandas as pd

def get_feature_importance(model, features):
    importance = model.get_feature_importance()
    df = pd.DataFrame({"Feature": features, "Importance": importance})
    df = df.sort_values(by="Importance", ascending=False)
    return df
