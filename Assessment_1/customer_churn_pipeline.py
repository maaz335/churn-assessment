import hashlib
import json
import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.base import BaseEstimator, TransformerMixin
from sklearn.cluster import KMeans
from sklearn.compose import ColumnTransformer
from sklearn.decomposition import PCA
from sklearn.ensemble import RandomForestClassifier
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
    roc_auc_score,
    roc_curve,
    silhouette_score,
)
from sklearn.model_selection import GridSearchCV, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

# student name
STUDENT_NAME = "Abdullah Aijaz"


# custom transformer to cap outliers using the IQR method
class TrainFittedIQRCapper(BaseEstimator, TransformerMixin):

    def __init__(self, factor=1.5):
        self.factor = factor
        self.bounds_ = {}

    # calculate lower and upper IQR boundaries on training data
    def fit(self, X, y=None):
        X_arr = np.asarray(X, dtype=float)
        self.bounds_ = {}
        for col_idx in range(X_arr.shape[1]):
            col_vals = X_arr[:, col_idx]
            valid_vals = col_vals[~np.isnan(col_vals)]
            if len(valid_vals) > 0:
                q1 = np.percentile(valid_vals, 25)
                q3 = np.percentile(valid_vals, 75)
                iqr = q3 - q1
                lower = q1 - self.factor * iqr
                upper = q3 + self.factor * iqr
                self.bounds_[col_idx] = (lower, upper)
            else:
                self.bounds_[col_idx] = (-np.inf, np.inf)
        return self

    # clip values between calculated lower and upper bounds
    def transform(self, X):
        X_arr = np.asarray(X, dtype=float).copy()
        for col_idx, (lower, upper) in self.bounds_.items():
            if col_idx < X_arr.shape[1]:
                X_arr[:, col_idx] = np.clip(X_arr[:, col_idx], lower, upper)
        return X_arr

    def get_feature_names_out(self, input_features=None):
        if input_features is not None:
            return np.asarray(input_features, dtype=object)
        return np.asarray([f"x{i}" for i in range(len(self.bounds_))], dtype=object)


# main machine learning pipeline class for telco customer churn
class CustomerChurnPipeline:

    # reference sha256 checksum for telco customer churn dataset
    EXPECTED_DATASET_HASH = "16320c9c1ec72448db59aa0a26a0b95401046bef5d02fd3aeb906448e3055e91"

    def __init__(self, data_path="customer_churn.csv", output_dir="outputs", random_state=42):
        self.data_path = data_path
        self.output_dir = output_dir
        self.random_state = random_state
        self.raw_data = None
        self.X_train = None
        self.X_test = None
        self.y_train = None
        self.y_test = None
        self.preprocessor = None
        self.trained_models = {}
        self.evaluation_results = {}
        self.unsupervised_results = {}

        # create output directory if it does not exist
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

    # verify dataset file exists and compare sha256 checksum
    def verify_data_integrity(self, expected_hash=None):
        if expected_hash is None:
            expected_hash = self.EXPECTED_DATASET_HASH

        if not os.path.isfile(self.data_path):
            raise FileNotFoundError(f"Dataset file not found at: {self.data_path}")

        hasher = hashlib.sha256()
        with open(self.data_path, "rb") as file:
            for chunk in iter(lambda: file.read(65536), b""):
                hasher.update(chunk)
        computed_hash = hasher.hexdigest()

        print(f"Data Checksum (SHA-256): {computed_hash}")
        matches = computed_hash.lower() == expected_hash.lower()
        if matches:
            print("Integrity Check: PASS (Checksum matches reference dataset).")
        else:
            print(f"Integrity Check: MISMATCH (Expected {expected_hash}, got {computed_hash}).")
            raise ValueError("Data integrity verification failed: Dataset checksum mismatch.")

        return computed_hash, matches

    # load raw dataset into pandas dataframe and validate columns
    def load_dataset(self):
        print("\nStep 1: Loading and Validating Dataset...")
        if not os.path.isfile(self.data_path):
            raise FileNotFoundError(f"File not found: {self.data_path}")

        self.raw_data = pd.read_csv(self.data_path)
        if self.raw_data.empty:
            raise ValueError(f"Dataset at {self.data_path} is empty.")

        # ensure essential columns exist in the dataframe
        expected_columns = {"customerID", "Churn", "tenure", "MonthlyCharges", "TotalCharges"}
        missing_expected = expected_columns - set(self.raw_data.columns)
        if missing_expected:
            raise ValueError(f"Dataset is missing essential columns: {missing_expected}")

        print(f"Dataset loaded: {self.raw_data.shape[0]} rows, {self.raw_data.shape[1]} columns.")
        return self.raw_data

    # clean raw columns and perform stratified train test split
    def clean_and_split_data(self, test_size=0.20):
        print("\nStep 2: Deterministic Schema Preparation and Stratified Train/Test Split...")
        df = self.raw_data.copy()

        # drop customer ID because it is just an arbitrary identifier
        if "customerID" in df.columns:
            df.drop(columns=["customerID"], inplace=True)
            print("Dropped identifier column: 'customerID'.")

        # convert TotalCharges to numeric and turn blank spaces into NaN
        df["TotalCharges"] = pd.to_numeric(df["TotalCharges"].astype(str).str.strip(), errors="coerce")
        nan_count = df["TotalCharges"].isna().sum()
        print(f"Detected {nan_count} blank entries in TotalCharges coerced to NaN (to be imputed on training data).")

        # map target variable Yes and No to binary 1 and 0
        if df["Churn"].dtype == object or str(df["Churn"].dtype).startswith("str"):
            target_mapping = {"Yes": 1, "No": 0}
            df["Churn"] = df["Churn"].map(target_mapping)

        valid_churn = df["Churn"].dropna()
        if not set(valid_churn.unique()).issubset({0, 1}):
            raise ValueError("Target 'Churn' must contain binary values (0/1 or 'No'/'Yes').")

        # plot churn class distribution
        plt.figure(figsize=(6, 4))
        sns.countplot(x=df["Churn"], hue=df["Churn"], palette=["steelblue", "crimson"], legend=False)
        plt.title("Customer Churn Distribution (0 = Retained, 1 = Churned)")
        plt.xlabel("Churn Status")
        plt.ylabel("Number of Customers")
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_dir, "churn_distribution.png"), dpi=300)
        plt.close()
        print("Saved plot: churn_distribution.png")

        # separate features and target
        X = df.drop(columns=["Churn"])
        y = df["Churn"]

        # 80/20 stratified split to preserve churn ratio in both sets
        self.X_train, self.X_test, self.y_train, self.y_test = train_test_split(
            X, y, test_size=test_size, random_state=self.random_state, stratify=y
        )

        print(f"Stratified partition created: Training set = {len(self.X_train)} rows, Test set = {len(self.X_test)} rows.")
        return self.X_train, self.X_test, self.y_train, self.y_test

    # build preprocessing pipelines for numeric and categorical columns
    def build_preprocessor(self):
        if self.X_train is None:
            raise ValueError("Data must be split using clean_and_split_data() before building preprocessor.")

        num_cols = ["tenure", "MonthlyCharges", "TotalCharges"]
        passthrough_cols = ["SeniorCitizen"] if "SeniorCitizen" in self.X_train.columns else []
        cat_cols = [
            c for c in self.X_train.columns
            if c not in num_cols and c not in passthrough_cols and (
                self.X_train[c].dtype == "object"
                or isinstance(self.X_train[c].dtype, (pd.CategoricalDtype, pd.StringDtype))
            )
        ]

        # numeric pipeline: median imputation, IQR outlier capping, and standard scaling
        num_pipeline = Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("capper", TrainFittedIQRCapper(factor=1.5)),
            ("scaler", StandardScaler()),
        ])

        # categorical pipeline: most frequent imputation and one hot encoding
        cat_pipeline = Pipeline([
            ("imputer", SimpleImputer(strategy="most_frequent")),
            ("encoder", OneHotEncoder(drop="first", handle_unknown="ignore", sparse_output=False)),
        ])

        # combine pipelines into a single ColumnTransformer
        transformers = [
            ("num", num_pipeline, num_cols),
            ("cat", cat_pipeline, cat_cols),
        ]
        if passthrough_cols:
            transformers.append(("pass", "passthrough", passthrough_cols))

        self.preprocessor = ColumnTransformer(
            transformers=transformers,
            remainder="drop",
        )
        return self.preprocessor

    # run unsupervised clustering with K-Means and visualize with PCA
    def run_unsupervised_models(self, selection_criterion="silhouette"):
        print("\nStep 3: Unsupervised Machine Learning (K-Means & PCA)...")
        if self.preprocessor is None:
            self.build_preprocessor()

        # fit preprocessor only on training data to prevent leakage
        X_train_proc = self.preprocessor.fit_transform(self.X_train)

        # test candidate cluster counts from 2 to 6
        k_range = range(2, 7)
        inertias = []
        silhouette_scores = []

        print("Evaluating candidate cluster counts:")
        for k in k_range:
            km = KMeans(n_clusters=k, random_state=self.random_state, n_init=10)
            labels = km.fit_predict(X_train_proc)
            score = silhouette_score(X_train_proc, labels)
            inertias.append(km.inertia_)
            silhouette_scores.append(score)
            print(f"  k={k}: Inertia = {km.inertia_:.2f}, Silhouette Score = {score:.4f}")

        # plot elbow curve and silhouette scores side by side
        fig, axes = plt.subplots(1, 2, figsize=(12, 4))
        axes[0].plot(k_range, inertias, marker="o", color="steelblue")
        axes[0].set_title("K-Means Elbow Method (Inertia)")
        axes[0].set_xlabel("Number of Clusters (k)")
        axes[0].set_ylabel("Inertia")
        axes[0].grid(True, linestyle="--", alpha=0.5)

        axes[1].plot(k_range, silhouette_scores, marker="s", color="seagreen")
        axes[1].set_title("Silhouette Score by Cluster Count")
        axes[1].set_xlabel("Number of Clusters (k)")
        axes[1].set_ylabel("Silhouette Score")
        axes[1].grid(True, linestyle="--", alpha=0.5)

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_dir, "kmeans_evaluation.png"), dpi=300)
        plt.close()
        print("Saved plot: kmeans_evaluation.png")

        # pick best k using highest silhouette score
        if selection_criterion == "silhouette":
            best_idx = int(np.argmax(silhouette_scores))
            selected_k = list(k_range)[best_idx]
            best_score = silhouette_scores[best_idx]
            rationale = (
                f"Selected k={selected_k} by maximizing Silhouette Score ({best_score:.4f}). "
                f"Note: k=3 also represents an elbow inflection point with inertia drop of "
                f"{inertias[0] - inertias[1]:.1f}."
            )
        else:
            selected_k = 3
            best_score = silhouette_scores[list(k_range).index(selected_k)]
            rationale = f"Selected k={selected_k} based on domain segmentation and elbow leveling."

        print(f"K-Means Selection: {rationale}")

        # fit final kmeans model with chosen k
        final_kmeans = KMeans(n_clusters=selected_k, random_state=self.random_state, n_init=10)
        cluster_labels = final_kmeans.fit_predict(X_train_proc)

        # print customer distribution across clusters
        unique_labels, counts = np.unique(cluster_labels, return_counts=True)
        cluster_dist = {int(u): int(c) for u, c in zip(unique_labels, counts)}
        print("Cluster distribution on training set:")
        for u, c in cluster_dist.items():
            print(f"  Cluster {u}: {c} customers ({c / len(cluster_labels) * 100:.1f}%)")

        # project features into 2D using PCA for visualization
        pca = PCA(n_components=2, random_state=self.random_state)
        X_pca = pca.fit_transform(X_train_proc)
        explained_var = pca.explained_variance_ratio_
        print(f"PCA projection: PC1 = {explained_var[0] * 100:.2f}% variance, PC2 = {explained_var[1] * 100:.2f}% variance.")

        # plot 2D customer clusters scatter plot
        plt.figure(figsize=(8, 6))
        scatter = plt.scatter(
            X_pca[:, 0], X_pca[:, 1],
            c=cluster_labels, cmap="viridis", alpha=0.6, edgecolors="none"
        )
        plt.title(f"Unsupervised Customer Clusters (k={selected_k}) Projected via PCA")
        plt.xlabel(f"Principal Component 1 ({explained_var[0] * 100:.1f}% explained variance)")
        plt.ylabel(f"Principal Component 2 ({explained_var[1] * 100:.1f}% explained variance)")
        plt.colorbar(scatter, label="Cluster Assignment")
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_dir, "pca_customer_clusters.png"), dpi=300)
        plt.close()
        print("Saved plot: pca_customer_clusters.png")

        self.unsupervised_results = {
            "candidate_k": list(k_range),
            "inertias": [float(i) for i in inertias],
            "silhouette_scores": [float(s) for s in silhouette_scores],
            "selected_k": selected_k,
            "selected_silhouette_score": float(best_score),
            "selection_rationale": rationale,
            "cluster_distribution": cluster_dist,
            "pca_explained_variance_ratio": [float(v) for v in explained_var],
        }
        return self.unsupervised_results

    # train and tune supervised models (Logistic Regression and Random Forest)
    def train_supervised_models(self):
        print("\nStep 4: Training and Optimizing Supervised ML Models...")
        if self.preprocessor is None:
            self.build_preprocessor()

        # model 1: Logistic Regression baseline pipeline
        print("Training Logistic Regression pipeline...")
        lr_pipeline = Pipeline([
            ("preprocessor", self.preprocessor),
            ("classifier", LogisticRegression(max_iter=1000, random_state=self.random_state)),
        ])
        lr_pipeline.fit(self.X_train, self.y_train)
        self.trained_models["Logistic Regression"] = lr_pipeline
        print("Logistic Regression pipeline trained successfully.")

        # model 2: Random Forest pipeline with GridSearchCV tuning
        print("Configuring Random Forest pipeline and executing GridSearchCV...")
        rf_pipeline = Pipeline([
            ("preprocessor", self.preprocessor),
            ("classifier", RandomForestClassifier(random_state=self.random_state)),
        ])

        # hyperparameter search space
        param_grid = {
            "classifier__n_estimators": [50, 100, 150],
            "classifier__max_depth": [6, 10, None],
            "classifier__min_samples_split": [2, 5],
        }

        # 3 fold cross validation scored by f1
        grid_search = GridSearchCV(
            estimator=rf_pipeline,
            param_grid=param_grid,
            cv=3,
            scoring="f1",
            n_jobs=-1,
        )
        grid_search.fit(self.X_train, self.y_train)

        best_rf_pipeline = grid_search.best_estimator_
        best_params = grid_search.best_params_
        print(f"Best Random Forest hyperparameters (scoring=f1): {best_params}")
        self.trained_models["Random Forest"] = best_rf_pipeline
        return self.trained_models

    # evaluate trained models on unseen test partition
    def evaluate_models(self):
        print("\nStep 5: Evaluating Models on Held-out Test Set...")
        if not self.trained_models:
            raise ValueError("Models must be trained before calling evaluate_models().")

        fig, axes = plt.subplots(1, 2, figsize=(11, 4))
        roc_fig = plt.figure(figsize=(7, 5))

        # evaluate each model on test set
        for idx, (name, model) in enumerate(self.trained_models.items()):
            y_pred = model.predict(self.X_test)
            y_prob = model.predict_proba(self.X_test)[:, 1]

            acc = accuracy_score(self.y_test, y_pred)
            prec = precision_score(self.y_test, y_pred, zero_division=0)
            rec = recall_score(self.y_test, y_pred, zero_division=0)
            f1 = f1_score(self.y_test, y_pred, zero_division=0)
            auc = roc_auc_score(self.y_test, y_prob)

            self.evaluation_results[name] = {
                "Accuracy": float(acc),
                "Precision": float(prec),
                "Recall": float(rec),
                "F1 Score": float(f1),
                "ROC AUC": float(auc),
            }

            print(f"\n{name} Held Out Evaluation")
            print(f"Accuracy:  {acc * 100:.2f}%")
            print(f"Precision: {prec * 100:.2f}%")
            print(f"Recall:    {rec * 100:.2f}%")
            print(f"F1 Score:  {f1 * 100:.2f}%")
            print(f"ROC AUC:   {auc * 100:.2f}%")
            print("Classification Report:\n", classification_report(self.y_test, y_pred, zero_division=0))

            # plot confusion matrix heatmap
            cm = confusion_matrix(self.y_test, y_pred)
            sns.heatmap(cm, annot=True, fmt="d", cmap="Blues", ax=axes[idx], cbar=False)
            axes[idx].set_title(f"Confusion Matrix: {name}")
            axes[idx].set_xlabel("Predicted Label")
            axes[idx].set_ylabel("True Label")
            axes[idx].set_xticklabels(["Stayed (0)", "Churned (1)"])
            axes[idx].set_yticklabels(["Stayed (0)", "Churned (1)"])

            # plot roc curve
            fpr, tpr, _ = roc_curve(self.y_test, y_prob)
            plt.figure(roc_fig.number)
            plt.plot(fpr, tpr, label=f"{name} (AUC = {auc:.3f})")

        fig.tight_layout()
        fig.savefig(os.path.join(self.output_dir, "confusion_matrices.png"), dpi=300)
        plt.close(fig)
        print("Saved plot: confusion_matrices.png")

        # finalize and save roc curves plot
        plt.figure(roc_fig.number)
        plt.plot([0, 1], [0, 1], "k--", label="Random Chance Baseline")
        plt.title("ROC Curves Comparison (Held-out Test Set)")
        plt.xlabel("False Positive Rate")
        plt.ylabel("True Positive Rate")
        plt.legend(loc="lower right")
        plt.grid(True, linestyle="--", alpha=0.5)
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_dir, "roc_curves.png"), dpi=300)
        plt.close(roc_fig)
        print("Saved plot: roc_curves.png")

        # extract and plot top 10 feature importances from random forest
        rf_pipe = self.trained_models["Random Forest"]
        fitted_preprocessor = rf_pipe.named_steps["preprocessor"]
        raw_feature_names = fitted_preprocessor.get_feature_names_out()
        clean_feature_names = [
            f.replace("num__", "").replace("cat__", "").replace("pass__", "")
            for f in raw_feature_names
        ]

        rf_clf = rf_pipe.named_steps["classifier"]
        importances = rf_clf.feature_importances_
        top_indices = np.argsort(importances)[-10:]

        plt.figure(figsize=(9, 5))
        plt.barh(range(len(top_indices)), importances[top_indices], align="center", color="darkslategray")
        plt.yticks(range(len(top_indices)), [clean_feature_names[i] for i in top_indices])
        plt.title("Top 10 Feature Importances (Random Forest)")
        plt.xlabel("Relative Importance Score")
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_dir, "feature_importance.png"), dpi=300)
        plt.close()
        print("Saved plot: feature_importance.png")

        # save all evaluation metrics to json file
        metrics_payload = {
            "supervised_models": self.evaluation_results,
            "unsupervised_models": self.unsupervised_results,
        }
        metrics_path = os.path.join(self.output_dir, "evaluation_metrics.json")
        with open(metrics_path, "w", encoding="utf-8") as f:
            json.dump(metrics_payload, f, indent=2)
        print(f"Saved evaluation metrics to: {metrics_path}")

        # print final summary table
        summary_df = pd.DataFrame(self.evaluation_results).T
        print("\n================ Model Benchmarking Summary ================")
        print(summary_df.round(4))
        print("============================================================\n")
        return self.evaluation_results

    # run complete end to end pipeline
    def run_full_pipeline(self):
        print("Starting Telco Customer Churn Pipeline Execution...")
        self.verify_data_integrity()
        self.load_dataset()
        self.clean_and_split_data()
        self.build_preprocessor()
        self.run_unsupervised_models()
        self.train_supervised_models()
        self.evaluate_models()
        print(f"All pipeline steps completed successfully! Output charts saved in: {self.output_dir}")


# run script directly from terminal
if __name__ == "__main__":
    pipeline = CustomerChurnPipeline(
        data_path="customer_churn.csv",
        output_dir="outputs",
        random_state=42,
    )
    pipeline.run_full_pipeline()
