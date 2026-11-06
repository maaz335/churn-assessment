"""
Automated Test Suite for CMP6221 Assessment 1 Reference Implementation
Covers 15 functional, leakage prevention, compatibility, and verification requirements.
"""

import json
import os
import shutil
import tempfile
import sys
import unittest
import numpy as np
import pandas as pd
from sklearn.pipeline import Pipeline

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from customer_churn_pipeline import CustomerChurnPipeline, TrainFittedIQRCapper


class TestCustomerChurnPipeline(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.test_dir = tempfile.mkdtemp(prefix="test_pipeline_")
        cls.output_dir = os.path.join(cls.test_dir, "outputs")
        cls.data_path = os.path.join(current_dir, "customer_churn.csv")
        if not os.path.exists(cls.data_path):
            cls.data_path = "customer_churn.csv"

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.test_dir, ignore_errors=True)

    def setUp(self):
        self.pipeline = CustomerChurnPipeline(
            data_path=self.data_path,
            output_dir=self.output_dir,
            random_state=42
        )

    # 1. Valid CSV loading
    def test_01_valid_csv_loading(self):
        raw = self.pipeline.load_dataset()
        self.assertIsInstance(raw, pd.DataFrame)
        self.assertFalse(raw.empty)
        self.assertEqual(raw.shape[0], 7043)
        self.assertIn("customerID", raw.columns)
        self.assertIn("Churn", raw.columns)

    # 2. Missing TotalCharges handling
    def test_02_missing_total_charges_handling(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.assertTrue(self.pipeline.X_train["TotalCharges"].dtype in [float, np.float64])
        total_nans = (
            self.pipeline.X_train["TotalCharges"].isna().sum() +
            self.pipeline.X_test["TotalCharges"].isna().sum()
        )
        self.assertEqual(total_nans, 11)

    # 3. Incorrect column detection / validation
    def test_03_incorrect_column_detection(self):
        bad_csv = os.path.join(self.test_dir, "bad.csv")
        pd.DataFrame({"colA": [1, 2], "colB": [3, 4]}).to_csv(bad_csv, index=False)
        bad_pipe = CustomerChurnPipeline(data_path=bad_csv, output_dir=self.output_dir)
        with self.assertRaises(ValueError):
            bad_pipe.load_dataset()

    # 4. Target encoding
    def test_04_target_encoding(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.assertTrue(set(self.pipeline.y_train.unique()).issubset({0, 1}))
        self.assertTrue(set(self.pipeline.y_test.unique()).issubset({0, 1}))

    # 5. Train/test split correctness & stratification
    def test_05_train_test_split_stratification(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data(test_size=0.20)
        total_rows = len(self.pipeline.X_train) + len(self.pipeline.X_test)
        self.assertEqual(total_rows, 7043)
        self.assertEqual(len(self.pipeline.X_test), 1409)

        train_rate = self.pipeline.y_train.mean()
        test_rate = self.pipeline.y_test.mean()
        self.assertAlmostEqual(train_rate, test_rate, delta=0.01)

    # 6. Prevention of training/test leakage
    def test_06_prevention_of_leakage(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        preprocessor = self.pipeline.build_preprocessor()

        preprocessor.fit(self.pipeline.X_train)
        orig_median = preprocessor.named_transformers_["num"].named_steps["imputer"].statistics_[2]
        orig_mean = preprocessor.named_transformers_["num"].named_steps["scaler"].mean_[2]

        corrupted_X_test = self.pipeline.X_test.copy()
        corrupted_X_test["TotalCharges"] = corrupted_X_test["TotalCharges"] * 1000.0

        transformed_test = preprocessor.transform(corrupted_X_test)

        current_median = preprocessor.named_transformers_["num"].named_steps["imputer"].statistics_[2]
        current_mean = preprocessor.named_transformers_["num"].named_steps["scaler"].mean_[2]

        self.assertEqual(orig_median, current_median)
        self.assertEqual(orig_mean, current_mean)

    # 7. Prevention of preprocessing leakage during GridSearchCV
    def test_07_gridsearch_pipeline_structure(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.pipeline.build_preprocessor()
        self.pipeline.train_supervised_models()

        rf_pipeline = self.pipeline.trained_models["Random Forest"]
        self.assertIsInstance(rf_pipeline, Pipeline)
        self.assertEqual(rf_pipeline.steps[0][0], "preprocessor")
        self.assertEqual(rf_pipeline.steps[1][0], "classifier")

    # 8. Unseen categorical values
    def test_08_unseen_categorical_values(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        preprocessor = self.pipeline.build_preprocessor()
        preprocessor.fit(self.pipeline.X_train)

        unseen_test = self.pipeline.X_test.iloc[:1].copy()
        unseen_test["PaymentMethod"] = "CryptoCurrency"
        unseen_test["Contract"] = "Decade-to-Decade"

        transformed = preprocessor.transform(unseen_test)
        self.assertEqual(transformed.shape[0], 1)
        self.assertEqual(transformed.shape[1], 30)

    # 9. Valid Logistic Regression predictions
    def test_09_valid_logistic_regression_predictions(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.pipeline.train_supervised_models()

        lr = self.pipeline.trained_models["Logistic Regression"]
        preds = lr.predict(self.pipeline.X_test)
        probs = lr.predict_proba(self.pipeline.X_test)

        self.assertEqual(len(preds), len(self.pipeline.X_test))
        self.assertTrue(set(preds).issubset({0, 1}))
        self.assertTrue(np.all(probs >= 0.0) and np.all(probs <= 1.0))
        self.assertAlmostEqual(np.sum(probs[0]), 1.0, places=4)

    # 10. Valid Random Forest predictions
    def test_10_valid_random_forest_predictions(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.pipeline.train_supervised_models()

        rf = self.pipeline.trained_models["Random Forest"]
        preds = rf.predict(self.pipeline.X_test)
        probs = rf.predict_proba(self.pipeline.X_test)

        self.assertEqual(len(preds), len(self.pipeline.X_test))
        self.assertTrue(set(preds).issubset({0, 1}))
        self.assertTrue(np.all(probs >= 0.0) and np.all(probs <= 1.0))

    # 11. K-Means without target features
    def test_11_kmeans_without_target_feature(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.assertNotIn("Churn", self.pipeline.X_train.columns)
        self.pipeline.run_unsupervised_models()
        proc_features = self.pipeline.preprocessor.transform(self.pipeline.X_train)
        self.assertEqual(proc_features.shape[1], 30)

    # 12. Cluster selection logic
    def test_12_cluster_selection_logic(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        res = self.pipeline.run_unsupervised_models(selection_criterion="silhouette")
        self.assertIn("selected_k", res)
        self.assertIn(res["selected_k"], range(2, 7))
        self.assertIn("selected_silhouette_score", res)
        self.assertEqual(res["selected_k"], 2)
        self.assertGreater(res["selected_silhouette_score"], 0.20)

    # 13. Model evaluation metrics
    def test_13_model_evaluation(self):
        self.pipeline.load_dataset()
        self.pipeline.clean_and_split_data()
        self.pipeline.train_supervised_models()
        metrics = self.pipeline.evaluate_models()

        for model_name in ["Logistic Regression", "Random Forest"]:
            self.assertIn(model_name, metrics)
            m = metrics[model_name]
            self.assertGreaterEqual(m["Accuracy"], 0.70)
            self.assertLessEqual(m["Accuracy"], 1.00)
            self.assertGreaterEqual(m["ROC AUC"], 0.75)
            self.assertLessEqual(m["ROC AUC"], 1.00)
            self.assertGreaterEqual(m["F1 Score"], 0.50)
            self.assertLessEqual(m["F1 Score"], 1.00)

    # 14. Expected chart generation
    def test_14_expected_chart_generation(self):
        self.pipeline.run_full_pipeline()
        expected_charts = [
            "churn_distribution.png",
            "confusion_matrices.png",
            "feature_importance.png",
            "kmeans_evaluation.png",
            "pca_customer_clusters.png",
            "roc_curves.png"
        ]
        for chart in expected_charts:
            chart_path = os.path.join(self.output_dir, chart)
            self.assertTrue(os.path.isfile(chart_path), f"Missing chart: {chart}")
            self.assertGreater(os.path.getsize(chart_path), 1000)

    # 15. Successful end-to-end execution
    def test_15_end_to_end_execution(self):
        new_pipe = CustomerChurnPipeline(
            data_path=self.data_path,
            output_dir=os.path.join(self.test_dir, "e2e_outputs"),
            random_state=42
        )
        new_pipe.run_full_pipeline()
        metrics_file = os.path.join(self.test_dir, "e2e_outputs", "evaluation_metrics.json")
        self.assertTrue(os.path.isfile(metrics_file))
        with open(metrics_file, "r") as f:
            data = json.load(f)
        self.assertIn("supervised_models", data)
        self.assertIn("unsupervised_models", data)


if __name__ == "__main__":
    unittest.main()
