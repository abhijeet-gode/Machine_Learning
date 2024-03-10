import joblib
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)
CORS(app)  # Enable CORS for all origins

# Load the trained XGBoost model
xgboost_model = joblib.load('XGBoost_model.pkl')

@app.route('/predict-wine', methods=['POST'])
def predict():
    try:
        data = request.form.to_dict()
        print("Received data:", data)  # Check if data is received correctly
        input_data = pd.DataFrame([data])
        input_data = input_data.astype(float)
        prediction = xgboost_model.predict(input_data)
        return jsonify({'prediction': prediction.tolist()})
    except Exception as e:
        print("Error:", e)  # Print out any exceptions for debugging
        return jsonify({'error': str(e)})

# Run the app
if __name__ == '__main__':
    app.run(debug=True)
