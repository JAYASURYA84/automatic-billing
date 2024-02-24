import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA

# Load the time series data of product demand
data = pd.read_csv('product_demand_data.csv', index_col='date', parse_dates=True)

# Data preprocessing (if needed)
# e.g., handle missing values, convert data types, resample, etc.

# Split the data into training and testing sets
train_data = data.iloc[:80]
test_data = data.iloc[80:]

# ARIMA model configuration
order = (1, 0, 0) # ARIMA(p, d, q) parameters

# Create and train ARIMA model
model = ARIMA(train_data, order=order)
model_fit = model.fit()

# Make forecasts
forecast_values = model_fit.forecast(steps=len(test_data))

# Model evaluation
mse = np.mean((forecast_values - test_data.values)**2)
rmse = np.sqrt(mse)
mae = np.mean(np.abs(forecast_values - test_data.values))

# Visualize the forecasts
plt.plot(train_data, label='Train Data')
plt.plot(test_data, label='Test Data')
plt.plot(test_data.index, forecast_values, label='Forecasted Values')
plt.xlabel('Date')
plt.ylabel('Product Demand')
plt.title('ARIMA Forecast for Product Demand')
plt.legend()
plt.show()
