"""K-Means clustering for humanitarian-aid screening.

Run from this folder with:
    python kmeans_humanitarian.py --data Country-data.csv --k 4

The script keeps country names out of the feature matrix, standardizes all
numeric variables, evaluates k=2..10, fits the selected K-Means model, writes
cluster outputs, and produces figures for the report.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import (
    calinski_harabasz_score,
    davies_bouldin_score,
    silhouette_score,
)
from sklearn.preprocessing import StandardScaler


FEATURES = [
    "child_mort",
    "exports",
    "health",
    "imports",
    "income",
    "inflation",
    "life_expec",
    "total_fer",
    "gdpp",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", type=Path, default=Path("Country-data.csv"))
    parser.add_argument("--outdir", type=Path, default=Path("results"))
    parser.add_argument("--figdir", type=Path, default=Path("figures"))
    parser.add_argument("--k", type=int, default=4, help="Selected number of clusters")
    parser.add_argument("--seed", type=int, default=42)
    return parser.parse_args()


def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in ["country", *FEATURES] if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    if df[FEATURES].isna().any().any():
        raise ValueError("The feature matrix contains missing values.")
    if df[FEATURES].duplicated().any():
        print("Warning: duplicated feature rows exist; they are retained.")
    return df


def evaluate_k_values(z: np.ndarray, seed: int) -> pd.DataFrame:
    rows = []
    for k in range(2, 11):
        model = KMeans(n_clusters=k, n_init=50, random_state=seed)
        labels = model.fit_predict(z)
        rows.append(
            {
                "k": k,
                "inertia": model.inertia_,
                "silhouette": silhouette_score(z, labels),
                "calinski_harabasz": calinski_harabasz_score(z, labels),
                "davies_bouldin": davies_bouldin_score(z, labels),
                "smallest_cluster": int(pd.Series(labels).value_counts().min()),
            }
        )
    return pd.DataFrame(rows)


def minmax_0_100(values: pd.Series) -> pd.Series:
    lo, hi = values.min(), values.max()
    if np.isclose(lo, hi):
        return pd.Series(50.0, index=values.index)
    return 100 * (values - lo) / (hi - lo)


def build_priority_outputs(
    df: pd.DataFrame,
    z: np.ndarray,
    labels: np.ndarray,
    scaler: StandardScaler,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Create a transparent, heuristic screening score in addition to clusters.

    The cluster assignment remains the K-Means result. The score is only a
    within-data screening aid, because the dataset has no conflict, disaster,
    displacement, food-security, or access-to-services variables.
    """
    scored = df[["country"]].copy()
    scored["cluster"] = labels
    scored["child_mort_z"] = z[:, FEATURES.index("child_mort")]
    scored["total_fer_z"] = z[:, FEATURES.index("total_fer")]
    scored["inflation_z"] = z[:, FEATURES.index("inflation")]
    scored["income_z"] = z[:, FEATURES.index("income")]
    scored["life_expec_z"] = z[:, FEATURES.index("life_expec")]
    scored["gdpp_z"] = z[:, FEATURES.index("gdpp")]
    scored["health_z"] = z[:, FEATURES.index("health")]

    # Higher score means greater socioeconomic vulnerability within this data.
    scored["screening_score_raw"] = (
        scored["child_mort_z"]
        + scored["total_fer_z"]
        + 0.5 * scored["inflation_z"]
        - scored["income_z"]
        - scored["life_expec_z"]
        - scored["gdpp_z"]
        - 0.5 * scored["health_z"]
    )
    scored["screening_score_0_100"] = minmax_0_100(scored["screening_score_raw"])

    profile = (
        df.assign(cluster=labels)
        .groupby("cluster", as_index=True)[FEATURES]
        .mean()
    )
    profile_z = pd.DataFrame(
        scaler.transform(profile[FEATURES]), index=profile.index, columns=FEATURES
    )
    profile_z["cluster_screening_score"] = (
        profile_z["child_mort"]
        + profile_z["total_fer"]
        + 0.5 * profile_z["inflation"]
        - profile_z["income"]
        - profile_z["life_expec"]
        - profile_z["gdpp"]
        - 0.5 * profile_z["health"]
    )
    risk_order = profile_z["cluster_screening_score"].sort_values(ascending=False)
    rank_map = {cluster: rank + 1 for rank, cluster in enumerate(risk_order.index)}
    scored["cluster_priority_rank"] = scored["cluster"].map(rank_map).astype(int)
    scored = scored.sort_values(
        ["cluster_priority_rank", "screening_score_0_100"], ascending=[True, False]
    )

    profile_out = profile.copy()
    profile_out["n_countries"] = df.assign(cluster=labels).groupby("cluster").size()
    profile_out["cluster_screening_score"] = profile_z["cluster_screening_score"]
    profile_out["cluster_priority_rank"] = profile_out.index.map(rank_map).astype(int)
    return scored, profile_out.sort_values("cluster_priority_rank")


def save_figures(
    eval_df: pd.DataFrame,
    z: np.ndarray,
    labels: np.ndarray,
    features: list[str],
    figdir: Path,
) -> None:
    figdir.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid", context="notebook")

    fig, axes = plt.subplots(1, 2, figsize=(12, 4.8), constrained_layout=True)
    axes[0].plot(eval_df["k"], eval_df["inertia"], marker="o", color="#1f4e79")
    axes[0].axvline(4, color="#c0504d", linestyle="--", linewidth=1.2, label="k = 4")
    axes[0].set_title("Elbow method")
    axes[0].set_xlabel("Number of clusters k")
    axes[0].set_ylabel("Within-cluster SSE")
    axes[0].legend(frameon=False)
    axes[1].plot(eval_df["k"], eval_df["silhouette"], marker="o", color="#2e8b57")
    axes[1].axvline(4, color="#c0504d", linestyle="--", linewidth=1.2, label="k = 4")
    axes[1].set_title("Silhouette score")
    axes[1].set_xlabel("Number of clusters k")
    axes[1].set_ylabel("Mean silhouette")
    axes[1].legend(frameon=False)
    fig.savefig(figdir / "01_k_selection.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    pca = PCA(n_components=2, random_state=0)
    coords = pca.fit_transform(z)
    fig, ax = plt.subplots(figsize=(8.2, 6.0), constrained_layout=True)
    palette = sns.color_palette("Set2", n_colors=len(np.unique(labels)))
    for cluster in sorted(np.unique(labels)):
        mask = labels == cluster
        ax.scatter(
            coords[mask, 0],
            coords[mask, 1],
            s=42,
            alpha=0.85,
            color=palette[cluster],
            label=f"Cụm {cluster}",
            edgecolor="white",
            linewidth=0.4,
        )
    ax.set_title(
        "Chiếu PCA hai chiều cho kết quả K-Means "
        f"(PC1 {pca.explained_variance_ratio_[0]:.1%}, "
        f"PC2 {pca.explained_variance_ratio_[1]:.1%})"
    )
    ax.set_xlabel("Thành phần chính PC1")
    ax.set_ylabel("Thành phần chính PC2")
    ax.legend(frameon=False, ncol=2)
    fig.savefig(figdir / "02_pca_clusters.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    profile_z = pd.DataFrame(z, columns=features).assign(cluster=labels).groupby("cluster").mean()
    fig, ax = plt.subplots(figsize=(11, 3.8), constrained_layout=True)
    sns.heatmap(
        profile_z,
        cmap="RdBu_r",
        center=0,
        annot=True,
        fmt=".1f",
        linewidths=0.5,
        cbar_kws={"label": "Trung bình chuẩn hóa (Z-score)"},
        ax=ax,
    )
    ax.set_title("Hồ sơ trung bình chuẩn hóa theo cụm")
    ax.set_xlabel("Biến đầu vào")
    ax.set_ylabel("Cụm")
    fig.savefig(figdir / "03_cluster_profile_heatmap.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    args = parse_args()
    if args.k < 2:
        raise ValueError("--k must be at least 2")
    df = load_data(args.data)
    X = df[FEATURES].astype(float)
    scaler = StandardScaler()
    z = scaler.fit_transform(X)

    eval_df = evaluate_k_values(z, args.seed)
    if args.k > len(df) - 1:
        raise ValueError("--k is too large for the number of rows")
    model = KMeans(n_clusters=args.k, n_init=50, random_state=args.seed)
    labels = model.fit_predict(z)

    args.outdir.mkdir(parents=True, exist_ok=True)
    args.figdir.mkdir(parents=True, exist_ok=True)
    eval_df.to_csv(args.outdir / "k_evaluation.csv", index=False)

    result = df[["country"] + FEATURES].copy()
    result["cluster"] = labels
    result.to_csv(args.outdir / "country_clusters.csv", index=False)

    country_scores, profile = build_priority_outputs(df, z, labels, scaler)
    country_scores.to_csv(args.outdir / "aid_screening.csv", index=False)
    profile.to_csv(args.outdir / "cluster_profile.csv")

    centers = pd.DataFrame(
        scaler.inverse_transform(model.cluster_centers_), columns=FEATURES
    )
    centers.index.name = "cluster"
    centers.to_csv(args.outdir / "cluster_centers_original_scale.csv")

    pca = PCA(n_components=2, random_state=0)
    coords = pca.fit_transform(z)
    pd.DataFrame(coords, columns=["PC1", "PC2"]).assign(
        country=df["country"].values, cluster=labels
    ).to_csv(args.outdir / "pca_coordinates.csv", index=False)
    save_figures(eval_df, z, labels, FEATURES, args.figdir)

    summary = {
        "n_countries": int(len(df)),
        "n_features": len(FEATURES),
        "selected_k": args.k,
        "seed": args.seed,
        "n_init": 50,
        "inertia": float(model.inertia_),
        "silhouette": float(silhouette_score(z, labels)),
        "calinski_harabasz": float(calinski_harabasz_score(z, labels)),
        "davies_bouldin": float(davies_bouldin_score(z, labels)),
        "pca_variance_ratio": pca.explained_variance_ratio_.tolist(),
        "cluster_sizes": {
            str(k): int(v) for k, v in pd.Series(labels).value_counts().sort_index().items()
        },
    }
    (args.outdir / "run_summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
