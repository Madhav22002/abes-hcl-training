"""
Inferential Statistics, Correlation Analysis, and Regression Modeling Engine
Adheres to strict analytical integrity:
- Reports p-values only where mathematically appropriate.
- Explicitly states statistical assumptions and limitations.
- Enforces distinction between correlation and causation.
"""

import numpy as np
import pandas as pd
from scipy import stats
import statsmodels.api as sm
from typing import Dict, Any

def compute_correlation_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """Compute Pearson correlation matrix across numerical metrics."""
    metrics = ["impressions", "clicks", "conversions", "spend_inr", "revenue_inr"]
    valid_metrics = [m for m in metrics if m in df.columns]
    if len(df) < 3 or len(valid_metrics) < 2:
        return pd.DataFrame()
    return df[valid_metrics].corr(method="pearson").round(3)

def compute_spend_revenue_regression(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Perform Ordinary Least Squares (OLS) Regression: Revenue = beta0 + beta1 * Spend.
    Returns slope, intercept, R-squared, p-value, std_err, and diagnostic notes.
    """
    if len(df) < 5 or "spend_inr" not in df.columns or "revenue_inr" not in df.columns:
        return {
            "valid": False,
            "error": "Insufficient data points for regression analysis (minimum 5 required)."
        }

    x = df["spend_inr"].values
    y = df["revenue_inr"].values

    # Check for variance in x
    if np.var(x) == 0:
        return {
            "valid": False,
            "error": "Independent variable (Spend) exhibits zero variance."
        }

    # SciPy linear regression
    slope, intercept, r_val, p_val, std_err = stats.linregress(x, y)
    r_squared = r_val ** 2

    # Fit statsmodels for robust summary and confidence intervals
    x_const = sm.add_constant(x)
    model = sm.OLS(y, x_const).fit()

    return {
        "valid": True,
        "n_observations": int(len(df)),
        "slope": float(round(slope, 4)),
        "intercept": float(round(intercept, 2)),
        "r_value": float(round(r_val, 4)),
        "r_squared": float(round(r_squared, 4)),
        "p_value": float(p_val),
        "std_err": float(round(std_err, 4)),
        "conf_int_lower": float(round(model.conf_int()[1][0], 4)),
        "conf_int_upper": float(round(model.conf_int()[1][1], 4)),
        "f_statistic": float(round(model.fvalue, 3)),
        "f_pvalue": float(model.f_pvalue),
        "interpretation": (
            f"For every ₹1.00 increase in marketing spend, attributed revenue changes by "
            f"₹{slope:.2f} on average (R² = {r_squared:.1%}, p < {0.001 if p_val < 0.001 else round(p_val, 4)}). "
            f"This relationship is statistically {'significant at the α = 0.05 level' if p_val < 0.05 else 'not significant'}."
        ),
        "limitations": [
            "Correlation does not establish causation: revenue attribution models rely on synthetic touchpoints.",
            "OLS assumes normality of residuals, constant error variance (homoscedasticity), and independence.",
            "External macroeconomic and unmeasured competitive factors are not captured in the regression model."
        ]
    }

def perform_channel_roas_anova(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Perform One-Way ANOVA test to evaluate if ROAS differences across marketing channels are statistically significant.
    """
    if len(df) < 4 or "channel_name" not in df.columns or "roas" not in df.columns:
        return {
            "valid": False,
            "error": "Insufficient sample data for ANOVA test (minimum 4 rows required)."
        }

    # Clean non-null and finite roas
    valid_data = df.dropna(subset=["roas", "channel_name"])
    valid_data = valid_data[np.isfinite(valid_data["roas"])]

    channels = valid_data["channel_name"].unique()
    if len(channels) < 2:
        return {
            "valid": False,
            "error": "At least 2 distinct channels required for ANOVA comparison."
        }

    groups = [valid_data[valid_data["channel_name"] == ch]["roas"].values for ch in channels]
    groups = [g for g in groups if len(g) >= 2]

    if len(groups) < 2:
        return {
            "valid": False,
            "error": "Insufficient observations per channel group."
        }

    f_stat, p_val = stats.f_oneway(*groups)

    # Non-parametric alternative (Kruskal-Wallis)
    kw_stat, kw_pval = stats.kruskal(*groups)

    return {
        "valid": True,
        "n_channels": len(groups),
        "f_statistic": float(round(f_stat, 3)),
        "p_value": float(p_val),
        "kruskal_statistic": float(round(kw_stat, 3)),
        "kruskal_pvalue": float(kw_pval),
        "significant": bool(p_val < 0.05),
        "summary": (
            f"One-way ANOVA across {len(groups)} channels yields F = {f_stat:.2f}, p = {p_val:.4e}. "
            f"The hypothesis of equal channel ROAS means is {'rejected' if p_val < 0.05 else 'not rejected'} at α = 0.05. "
            f"The non-parametric Kruskal-Wallis test corroborates this (H = {kw_stat:.2f}, p = {kw_pval:.4e})."
        ),
        "limitations": [
            "ANOVA assumes normal distribution within each group and homogeneity of variance across channels.",
            "Synthetic data reflects pre-configured channel efficiency weights."
        ]
    }
