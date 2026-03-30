from flask import Flask, render_template, request
import pandas as pd
from catboost import CatBoostRegressor
from utils.preprocessing import preprocess_input
from utils.variance import calculate_variance
from utils.feature_importance import get_feature_importance

app = Flask(__name__)

# Categorical feature list
categorical_features = [
    'Material', 'Color', 'Finish', 'Water Resistancy',
    'Shape', 'Crystal Shape', 'Crystal Material'
]

# Numeric feature list
numeric_features = [
    'Size (12-6)', 'Size (3-9)', 'Height',
    'Crown Groove Diameter', 'Lug Width',
    'Dial VO 6H - 12H', 'Dial VO 3H - 9H'
]

# Sample dropdown options used when Main_Data.xlsx is not present
SAMPLE_OPTIONS = {
    'Material':         ['Brass', 'Stainless Steel', 'Titanium', 'Zinc Alloy'],
    'Color':            ['Black', 'Gold', 'Rose Gold', 'Silver'],
    'Finish':           ['Brushed', 'Matte', 'PVD Coated', 'Polished'],
    'Water Resistancy': ['3 ATM', '5 ATM', '10 ATM', '30 ATM'],
    'Shape':            ['Cushion', 'Oval', 'Rectangle', 'Round', 'Square'],
    'Crystal Shape':    ['Oval', 'Rectangle', 'Round', 'Square'],
    'Crystal Material': ['Acrylic', 'Mineral Glass', 'Sapphire'],
}

# Load model (optional — demo mode used if not found)
model = None
try:
    _m = CatBoostRegressor()
    _m.load_model("model/model_optimized_final.cbm")
    model = _m
except Exception:
    pass

# Load dataset for dropdowns (falls back to sample options if not found)
try:
    main_data = pd.read_excel("Main_Data.xlsx")
    numeric_features = [
        col for col in main_data.select_dtypes(include='number').columns
        if col != "Price" and col not in categorical_features
    ]
    dropdown_options = {
        col: sorted(main_data[col].dropna().unique())
        for col in categorical_features
    }
except Exception:
    dropdown_options = SAMPLE_OPTIONS


@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    error = None

    if request.method == "POST":
        if model is None:
            # Demo mode — simulate a realistic result for interface showcase
            import random
            demo_price = round(random.uniform(800, 2400), 2)
            variance_low, variance_high = calculate_variance(demo_price, 5)
            result = {
                "price": demo_price,
                "low": round(variance_low, 2),
                "high": round(variance_high, 2),
                "margin": 5,
                "demo": True
            }
        else:
            user_input = {}

            for col in categorical_features:
                user_input[col] = request.form.get(col)

            for col in numeric_features:
                val = request.form.get(col)
                user_input[col] = float(val) if val else 0.0

            input_df = pd.DataFrame([user_input])
            input_df = preprocess_input(input_df, categorical_features)

            prediction = model.predict(input_df)[0]

            variance_percent = 5
            variance_low, variance_high = calculate_variance(prediction, variance_percent)

            result = {
                "price": round(prediction, 2),
                "low": round(variance_low, 2),
                "high": round(variance_high, 2),
                "margin": variance_percent,
                "demo": False
            }

    return render_template(
        "index.html",
        options=dropdown_options,
        numeric_features=numeric_features,
        result=result,
        error=error,
        model_loaded=model is not None
    )

if __name__ == "__main__":
    app.run(debug=True)
