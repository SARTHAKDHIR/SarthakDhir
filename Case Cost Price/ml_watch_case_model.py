import pandas as pd
import numpy as np
from catboost import CatBoostRegressor, Pool
from sklearn.model_selection import KFold
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

# 1. Load dataset
data = pd.read_excel("Main_Data.xlsx")

# 2. Define features
categorical_features = ['Material', 'Color', 'Finish', 'Water Resistancy', 'Shape', 'Crystal Shape', 'Crystal Material']
numeric_features = ['Size (12-6)', 'Size (3-9)', 'Height', 'Crown Groove Diameter', 'Lug Width',
                    'Dial VO 6H - 12H', 'Dial VO 3H - 9H']
target = 'Price'

# 3. Prepare X and y
X = data[categorical_features + numeric_features].copy()
y = data[target].copy()

# 4. Ensure categorical columns are strings
for col in categorical_features:
    X[col] = X[col].astype(str)

# 5. Set up K-Fold
kf = KFold(n_splits=5, shuffle=True, random_state=42)

# 6. Initialize metric lists
rmse_list, mae_list, r2_list, mape_list = [], [], [], []

# 7. K-Fold Cross Validation
fold = 1
for train_index, val_index in kf.split(X):
    print(f"\n--- Fold {fold} ---")
    X_train, X_val = X.iloc[train_index], X.iloc[val_index]
    y_train, y_val = y.iloc[train_index], y.iloc[val_index]

    train_pool = Pool(data=X_train, label=y_train, cat_features=categorical_features)
    val_pool = Pool(data=X_val, label=y_val, cat_features=categorical_features)

    model = CatBoostRegressor(
        iterations=1000,
        learning_rate=0.05,
        depth=6,
        loss_function='RMSE',
        eval_metric='RMSE',
        random_seed=42,
        verbose=100
    )

    model.fit(train_pool, eval_set=val_pool, early_stopping_rounds=50)
    preds = model.predict(X_val)

    rmse = np.sqrt(mean_squared_error(y_val, preds))
    mae = mean_absolute_error(y_val, preds)
    r2 = r2_score(y_val, preds)
    mape = np.mean(np.abs((y_val - preds) / y_val)) * 100

    print(f"Fold {fold} — RMSE: {rmse:.2f} | MAE: {mae:.2f} | R2: {r2:.3f} | MAPE: {mape:.2f}%")

    rmse_list.append(rmse)
    mae_list.append(mae)
    r2_list.append(r2)
    mape_list.append(mape)
    fold += 1

# 8. Average metrics
print("\n=== Average Metrics Across 5 Folds ===")
print(f"RMSE : {np.mean(rmse_list):.2f} ± {np.std(rmse_list):.2f}")
print(f"MAE  : {np.mean(mae_list):.2f} ± {np.std(mae_list):.2f}")
print(f"R2   : {np.mean(r2_list):.3f} ± {np.std(r2_list):.3f}")
print(f"MAPE : {np.mean(mape_list):.2f}% ± {np.std(mape_list):.2f}%")

# 9. Train final model on full dataset
final_pool = Pool(data=X, label=y, cat_features=categorical_features)
final_model = CatBoostRegressor(
    iterations=1000,
    learning_rate=0.05,
    depth=6,
    loss_function='RMSE',
    eval_metric='RMSE',
    random_seed=42,
    verbose=100
)
final_model.fit(final_pool)

# 10. Feature importance
feature_importances = pd.DataFrame({
    'Feature': X.columns,
    'Importance': final_model.get_feature_importance()
}).sort_values(by='Importance', ascending=False)

print("\nFeature Importances:")
print(feature_importances)

# 11. Save the final model
final_model.save_model("model/model_optimized_final.cbm")
print("Model saved to model/model_optimized_final.cbm")
