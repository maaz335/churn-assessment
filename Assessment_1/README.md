CMP6221 Computing for AI Assessment 1
Educational Reference Implementation: Telco Customer Churn Prediction and Behavioral Segmentation

Institution: Birmingham City University
Module: CMP6221 Computing for AI
Assessment: Assessment 1 AI Based Computing Program
Academic Year: 2026 to 2027
Student Name: Abdullah Aijaz

Section 1: Project Objective and Real World Problem

Customer churn represents the rate at which subscribers discontinue their commercial relationship with a service provider. In the telecommunications sector, acquiring new customers costs between 5 to 25 times more than retaining existing subscribers. When a customer terminates service, the provider forfeits recurring monthly subscription revenue and onboarding capital investments.

This project delivers an end to end Machine Learning pipeline written in Python using an Object Oriented Programming architecture. The system performs:
1. Unsupervised Behavioral Clustering: Discovers natural customer personas without relying on ground truth target labels using K Means and Principal Component Analysis (PCA).
2. Supervised Churn Classification: Accurately identifies high risk churn candidates using Logistic Regression and Random Forest with GridSearchCV hyperparameter optimization.
3. Strict Data Leakage Prevention: Encapsulates all preprocessing (median imputation, outlier handling, one hot encoding, and standard scaling) within scikit learn Pipeline and ColumnTransformer constructs, ensuring zero information leaks across train and test splits or cross validation folds.

Section 2: Dataset Information and Acquisition

The project utilizes the industry benchmark Telco Customer Churn dataset originally published by IBM and Kaggle:
1. Observations: 7043 customer accounts
2. Attributes: 21 columns covering demographics, subscribed services, account tenure, payment channels, billing amounts, and churn target
3. File: customer_churn.csv
4. SHA256 Checksum Fingerprint: 16320c9c1ec72448db59aa0a26a0b95401046bef5d02fd3aeb906448e3055e91

Feature Categories:
1. Demographics: gender, SeniorCitizen, Partner, Dependents
2. Subscribed Services: PhoneService, MultipleLines, InternetService, OnlineSecurity, OnlineBackup, DeviceProtection, TechSupport, StreamingTV, StreamingMovies
3. Account and Billing: tenure, Contract, PaperlessBilling, PaymentMethod, MonthlyCharges, TotalCharges
4. Target Variable: Churn (Yes equals 1, No equals 0)

Section 3: System Requirements and Installation

Prerequisites:
Python 3.10, 3.11, or 3.12 with standard package manager pip.

Installation:
pip install -r requirements.txt

Dependencies:
numpy, pandas, scikit learn, scipy, matplotlib, seaborn, pytest

Section 4: How to Run the Application

Option A: Command Line Interface
Run the main script directly from the project folder:
python customer_churn_pipeline.py

Execution takes approximately 10 to 15 seconds and automatically:
1. Verifies dataset SHA256 checksum against reference and halts execution if a mismatch is detected.
2. Loads raw data and drops non predictive customerID.
3. Coerces 11 whitespace entries in TotalCharges to NaN.
4. Performs a stratified 80 20 train and test split.
5. Evaluates K Means clustering from k equals 2 to 6 and selects optimal cluster count.
6. Projects customer clusters in 2D space using PCA.
7. Trains Logistic Regression and optimizes Random Forest via 3 fold cross validated GridSearchCV.
8. Evaluates models on the unseen 20 percent test partition (1409 customers).
9. Generates all 6 publication grade figures in the outputs folder.
10. Saves JSON metrics to outputs/evaluation_metrics.json.

Option B: Interactive Jupyter Notebook
Launch Jupyter Notebook:
jupyter notebook customer_churn_notebook.ipynb

Option C: Run Automated Test Suite
Execute the 15 unit verification tests:
python test_pipeline.py

Section 5: Machine Learning and Preprocessing Architecture

Strict Data Leakage Prevention:
1. Stratified Split First: The 80 20 train and test split occurs before any transformer is fitted.
2. ColumnTransformer Construction:
Numerical Pipeline uses SimpleImputer with median strategy, TrainFittedIQRCapper with factor 1.5, and StandardScaler. All parameters are learned strictly from the training partition.
Categorical Pipeline uses SimpleImputer with most frequent strategy and OneHotEncoder with handle unknown set to ignore.
SeniorCitizen is passed through as a binary indicator.
3. Cross Validation Leakage Prevention: ColumnTransformer is nested directly inside Pipeline passed to GridSearchCV so preprocessing is fitted exclusively on training folds.

Machine Learning Models:
1. Unsupervised Model 1: K Means Clustering evaluates candidate clusters from k equals 2 to 6 using inertia and silhouette score. Dynamically selects k equals 2 by highest silhouette score 0.2952.
2. Unsupervised Model 2: Principal Component Analysis (PCA) reduces features into 2 orthogonal axes (PC1: 39.4 percent variance, PC2: 18.5 percent variance) to visualize customer clusters.
3. Supervised Model 1: Logistic Regression linear probabilistic classifier with L2 regularization.
4. Supervised Model 2: Random Forest Classifier ensemble model optimized via GridSearchCV over max depth, n estimators, and min samples split.

Section 6: Generated Output Artifacts

All charts are saved automatically to the outputs directory:
1. churn_distribution.png: Target class distribution (Retained vs Churned)
2. kmeans_evaluation.png: Side by side plot of Elbow inertia curve and Silhouette score curve
3. pca_customer_clusters.png: 2D PCA scatter plot showing assigned customer clusters
4. confusion_matrices.png: Side by side confusion matrices for Logistic Regression and Random Forest
5. roc_curves.png: Receiver Operating Characteristic curves with computed AUC values
6. feature_importance.png: Top 10 feature importances from Random Forest model
7. evaluation_metrics.json: Performance metrics file

Section 7: Assessment 1 Requirements Mapping

1. Python Program addressing a real world case study: PASS (customer_churn_pipeline.py implements end to end churn analysis)
2. Data Preprocessing and Cleaning: PASS (Dropped customerID, coerced blank TotalCharges, median imputation, train fitted IQR outlier handling, OneHotEncoder)
3. Programming Best Practices (OOP, Modular Functions, Clean Naming): PASS (CustomerChurnPipeline and TrainFittedIQRCapper classes)
4. At least 3 Machine Learning Algorithms: PASS (Logistic Regression, Random Forest, and K Means Clustering)
5. At least 1 Unsupervised Model: PASS (K Means Clustering and PCA)
6. Hyperparameter Optimization and Benchmarking: PASS (GridSearchCV on Random Forest with 3 fold CV and full comparison table)
7. Integrity and Checksum Verification: PASS (Cryptographic SHA256 fingerprinting that halts on mismatch)
8. Submission Format: PASS (Packaged into AbdullahAijaz.zip containing required python, test, and data files)

Section 8: Note on Moodle Specified AI Model Instruction

The assessment brief notes that the AI model will be specified on Moodle. Because direct Moodle access is not available, the project implements a standard linear baseline (Logistic Regression), a tuned non-linear ensemble (Random Forest with GridSearchCV), and an unsupervised clustering algorithm (K Means). The pipeline architecture is fully modular, allowing any specific model mandated by the lecturer to be substituted into the pipeline without altering the data processing workflow.

Section 9: Submission Archive

Assessment 1 submission package is ready as AbdullahAijaz.zip.
