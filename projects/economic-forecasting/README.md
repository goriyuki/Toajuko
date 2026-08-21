# Economic Indicator Forecasting

This project turns an MSc forecasting exercise into a small, reusable time-series workflow. The public version is deliberately data-agnostic: it does not include assessment prompts, student identifiers, submitted reports, or copied datasets.

## Problem

Monthly economic indicators often combine long-term trend, seasonality, structural shocks, and missing observations. The objective is to compare transparent ARIMA and SARIMA baselines on a held-out period, select the stronger model using forecast errors, and produce a future forecast with a 95% confidence interval.

## Workflow

1. Parse and validate a monthly time series from CSV.
2. Hold out the most recent observations for out-of-sample evaluation.
3. Fit non-seasonal ARIMA and seasonal ARIMA (SARIMA) candidates.
4. Compare MAE, RMSE, and MAPE; MAPE ignores zero-valued observations.
5. Refit the selected specification on all observations.
6. Export validation metrics, forecasts, and a plot.

## Run locally

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt

python src/forecasting.py \
  --data path/to/monthly_series.csv \
  --date-column Date \
  --value-column Value \
  --output outputs/forecast
```

Expected input:

```csv
Date,Value
2020-01-01,100.1
2020-02-01,100.4
```

Useful options include `--test-size`, `--forecast-steps`, `--arima-order`, and `--seasonal-order`. Run `python src/forecasting.py --help` for details.

## Outputs

- `metrics.csv`: hold-out errors for each model.
- `forecast.csv`: point forecast and confidence bounds for the selected model.
- `forecast.png`: observed history and forecast interval.

## Limitations

The default model orders are baselines, not universal choices. A production study should add rolling-origin validation, domain-informed shock treatment, alternative specifications, and data-revision checks. Forecasts are analytical outputs rather than financial advice.
