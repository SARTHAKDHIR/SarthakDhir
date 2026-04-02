# Case Cost Price AI

A production ML web application that predicts watch case prices based on physical specifications and material properties, built using CatBoost and Flask.

## Overview

This tool assists in estimating watch case costs during product development. Given a set of specifications, it predicts the case price and provides a ±5% variance range to account for supplier fluctuation.

> **Note:** The trained model (`.cbm`) and dataset (`.xlsx`) are proprietary and not included in this repository. The full training pipeline and application code are provided for reference.

## Features

- Predicts case price from material, dimensions, and finish specifications
- Displays a variance range with a price analysis chart
- Clean dark UI built with Flask + Chart.js
- CatBoost model trained with 5-fold cross-validation

## Tech Stack

| Layer    | Technology                        |
|----------|-----------------------------------|
| ML       | CatBoost, scikit-learn, pandas    |
| Backend  | Flask (Python)                    |
| Frontend | HTML, CSS, Chart.js               |

## Project Structure

```
├── app.py                        # Flask application
├── ml_watch_case_model.py        # Model training script (K-Fold + final model)
├── requirements.txt
├── utils/
│   ├── preprocessing.py          # Input preprocessing for CatBoost
│   ├── variance.py               # ±5% variance calculation
│   └── feature_importance.py     # Feature importance utility
├── templates/
│   └── index.html                # Web UI
├── static/css/
│   └── style.css
└── model/
    └── model_optimized_final.cbm # [NOT INCLUDED - proprietary]
```

## Model Details

- **Algorithm:** CatBoostRegressor
- **Validation:** 5-Fold Cross Validation
- **Categorical Features:** Material, Color, Finish, Water Resistancy, Shape, Crystal Shape, Crystal Material
- **Numeric Features:** Size (12-6), Size (3-9), Height, Crown Groove Diameter, Lug Width, Dial dimensions
- **Target:** Case Price (₹)

## Setup (with your own data)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Add your dataset
# Place Main_Data.xlsx in the root with the columns listed above + a "Price" column

# 3. Train the model
python ml_watch_case_model.py

# 4. Run the app
python app.py
```

Open `http://localhost:5000`

## Author

Developed by Sarthak Dhir
