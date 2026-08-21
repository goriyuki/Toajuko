"""Compare ARIMA and SARIMA forecasts for a monthly univariate series."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.statespace.sarimax import SARIMAX


@dataclass(frozen=True)
class ModelScore:
    model: str
    mae: float
    rmse: float
    mape: float


def parse_order(value: str, expected_length: int) -> tuple[int, ...]:
    """Parse a comma-separated integer model order."""
    try:
        order = tuple(int(part.strip()) for part in value.split(","))
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Model orders must contain integers") from exc
    if len(order) != expected_length or any(part < 0 for part in order):
        raise argparse.ArgumentTypeError(
            f"Expected {expected_length} non-negative comma-separated integers"
        )
    return order


def load_monthly_series(path: Path, date_column: str, value_column: str) -> pd.Series:
    """Load, validate, and regularise a monthly series."""
    frame = pd.read_csv(path, usecols=[date_column, value_column])
    frame[date_column] = pd.to_datetime(frame[date_column], errors="coerce")
    frame[value_column] = pd.to_numeric(frame[value_column], errors="coerce")
    frame = frame.dropna().sort_values(date_column).drop_duplicates(date_column)
    if frame.empty:
        raise ValueError("No valid date/value observations were found")

    series = frame.set_index(date_column)[value_column].asfreq("MS")
    if series.isna().any():
        raise ValueError(
            "The monthly index has missing observations; impute them explicitly before modelling"
        )
    if len(series) < 36:
        raise ValueError("At least 36 monthly observations are required")
    return series.astype(float)


def error_metrics(actual: pd.Series, predicted: pd.Series, model: str) -> ModelScore:
    """Calculate scale-dependent errors and zero-safe MAPE."""
    actual_values = actual.to_numpy(dtype=float)
    predicted_values = np.asarray(predicted, dtype=float)
    errors = actual_values - predicted_values
    non_zero = actual_values != 0
    mape = (
        float(np.mean(np.abs(errors[non_zero] / actual_values[non_zero])) * 100)
        if non_zero.any()
        else float("nan")
    )
    return ModelScore(
        model=model,
        mae=float(np.mean(np.abs(errors))),
        rmse=float(np.sqrt(np.mean(errors**2))),
        mape=mape,
    )


def evaluate_models(
    train: pd.Series,
    test: pd.Series,
    arima_order: tuple[int, int, int],
    seasonal_order: tuple[int, int, int, int],
) -> tuple[pd.DataFrame, str]:
    """Fit two baseline models and return their hold-out scores."""
    arima_fit = ARIMA(train, order=arima_order).fit()
    arima_forecast = arima_fit.forecast(steps=len(test))

    sarima_fit = SARIMAX(
        train,
        order=arima_order,
        seasonal_order=seasonal_order,
        enforce_stationarity=False,
        enforce_invertibility=False,
    ).fit(disp=False)
    sarima_forecast = sarima_fit.forecast(steps=len(test))

    scores = [
        error_metrics(test, arima_forecast, "ARIMA"),
        error_metrics(test, sarima_forecast, "SARIMA"),
    ]
    metrics = pd.DataFrame([score.__dict__ for score in scores]).sort_values("rmse")
    return metrics, str(metrics.iloc[0]["model"])


def forecast_selected_model(
    series: pd.Series,
    model_name: str,
    arima_order: tuple[int, int, int],
    seasonal_order: tuple[int, int, int, int],
    steps: int,
) -> pd.DataFrame:
    """Refit the selected model on all observations and forecast forward."""
    if model_name == "ARIMA":
        fitted = ARIMA(series, order=arima_order).fit()
    else:
        fitted = SARIMAX(
            series,
            order=arima_order,
            seasonal_order=seasonal_order,
            enforce_stationarity=False,
            enforce_invertibility=False,
        ).fit(disp=False)

    summary = fitted.get_forecast(steps=steps).summary_frame(alpha=0.05)
    return summary.rename(
        columns={
            "mean": "forecast",
            "mean_ci_lower": "lower_95",
            "mean_ci_upper": "upper_95",
        }
    )[["forecast", "lower_95", "upper_95"]]


def save_plot(series: pd.Series, forecast: pd.DataFrame, model_name: str, path: Path) -> None:
    """Save a compact forecast chart."""
    figure, axis = plt.subplots(figsize=(10, 5))
    axis.plot(series.index, series.values, label="Observed", color="#184E77")
    axis.plot(forecast.index, forecast["forecast"], label=model_name, color="#F28E2B")
    axis.fill_between(
        forecast.index,
        forecast["lower_95"],
        forecast["upper_95"],
        color="#F28E2B",
        alpha=0.2,
        label="95% confidence interval",
    )
    axis.set_title(f"Monthly forecast — selected model: {model_name}")
    axis.grid(alpha=0.2)
    axis.legend()
    figure.tight_layout()
    figure.savefig(path, dpi=160)
    plt.close(figure)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", required=True, type=Path, help="Input CSV path")
    parser.add_argument("--date-column", default="Date")
    parser.add_argument("--value-column", default="Value")
    parser.add_argument("--test-size", type=int, default=12)
    parser.add_argument("--forecast-steps", type=int, default=12)
    parser.add_argument("--arima-order", default="1,1,1")
    parser.add_argument("--seasonal-order", default="1,0,1,12")
    parser.add_argument("--output", type=Path, default=Path("outputs/forecast"))
    return parser


def main() -> None:
    args = build_parser().parse_args()
    if args.test_size < 1 or args.forecast_steps < 1:
        raise ValueError("Test size and forecast steps must be positive")

    arima_order = parse_order(args.arima_order, 3)
    seasonal_order = parse_order(args.seasonal_order, 4)
    series = load_monthly_series(args.data, args.date_column, args.value_column)
    if args.test_size >= len(series):
        raise ValueError("Test size must be smaller than the number of observations")

    train, test = series.iloc[: -args.test_size], series.iloc[-args.test_size :]
    metrics, selected_model = evaluate_models(
        train, test, arima_order, seasonal_order
    )
    forecast = forecast_selected_model(
        series,
        selected_model,
        arima_order,
        seasonal_order,
        args.forecast_steps,
    )

    args.output.mkdir(parents=True, exist_ok=True)
    metrics.to_csv(args.output / "metrics.csv", index=False)
    forecast.to_csv(args.output / "forecast.csv", index_label="Date")
    save_plot(series, forecast, selected_model, args.output / "forecast.png")
    print(metrics.to_string(index=False))
    print(f"Selected model: {selected_model}")
    print(f"Outputs written to: {args.output.resolve()}")


if __name__ == "__main__":
    main()
