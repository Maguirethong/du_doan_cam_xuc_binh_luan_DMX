import os
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.base import BaseEstimator, TransformerMixin
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.svm import LinearSVC
from sklearn.pipeline import Pipeline
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    classification_report,
    confusion_matrix
)
import joblib

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Lớp trích chọn đặc trưng Word Embeddings
class WordEmbeddingVectorizer(BaseEstimator, TransformerMixin):
    def __init__(self, vector_size=100, window=5, min_count=2):
        self.vector_size = vector_size
        self.window = window
        self.min_count = min_count

    def fit(self, X, y=None):
        from gensim.models import Word2Vec
        tokenized_corpus = [str(doc).split() for doc in X]
        self.model = Word2Vec(
            sentences=tokenized_corpus,
            vector_size=self.vector_size,
            window=self.window,
            min_count=self.min_count,
            workers=4
        )
        self.use_gensim = True
        return self

    def transform(self, X):
        if getattr(self, 'use_gensim', False):
            vectors = []
            for doc in X:
                words = [w for w in str(doc).split() if w in self.model.wv]
                if words:
                    vec = np.mean([self.model.wv[w] for w in words], axis=0)
                else:
                    vec = np.zeros(self.vector_size)
                vectors.append(vec)
            return np.array(vectors)
        else:
            tfidf_mat = self.tfidf.transform(X)
            return self.svd.transform(tfidf_mat)


# ĐỌC VÀ TIỀN XỬ LÝ DỮ LIỆU
def load_and_preprocess_data(
    filepath="comments_merged.xlsx",
    test_size=0.2,
    random_state=42
):
    df = pd.read_excel(filepath)

    df = df.dropna(subset=["Comment", "Nhãn"]).copy()

    for col in ["Comment", "Nhãn"]:
        df[col] = df[col].astype(str).str.strip()

    df = df[df["Comment"] != ""]

    X = df["Comment"]
    y = df["Nhãn"]

    return train_test_split(
        X, y,
        test_size=test_size,
        random_state=random_state,
        stratify=y
    )


# XÂY DỰNG CÁC PIPELINE 
def build_all_pipelines():
    pipelines = {
        # Naive Bayes với BoW và TF-IDF
        "BoW + Naive Bayes": Pipeline([
            ("vectorizer", CountVectorizer(ngram_range=(1, 2), max_features=10000)),
            ("classifier", MultinomialNB(alpha=1.0))
        ]),
        "TF-IDF + Naive Bayes": Pipeline([
            ("vectorizer", TfidfVectorizer(ngram_range=(1, 2), max_features=10000, sublinear_tf=True)),
            ("classifier", MultinomialNB(alpha=1.0))
        ]),

        # SVM với cả 3 phương pháp (BoW, TF-IDF, Word Embeddings)
        "BoW + SVM": Pipeline([
            ("vectorizer", CountVectorizer(ngram_range=(1, 2), max_features=10000)),
            ("classifier", LinearSVC(C=1.0, max_iter=3000, random_state=42))
        ]),
        "TF-IDF + SVM": Pipeline([
            ("vectorizer", TfidfVectorizer(ngram_range=(1, 2), max_features=10000, sublinear_tf=True)),
            ("classifier", LinearSVC(C=1.0, max_iter=3000, random_state=42))
        ]),
        "Word Embeddings + SVM": Pipeline([
            ("vectorizer", WordEmbeddingVectorizer(vector_size=100, window=5, min_count=2)),
            ("classifier", LinearSVC(C=1.0, max_iter=3000, random_state=42))
        ])
    }
    return pipelines


# ĐÁNH GIÁ MÔ HÌNH VÀ VẼ MA TRẬN NHẦM LẪN 
def evaluate_single_model(model_name, pipeline, X_test, y_test, output_dir="evaluation_results"):
    os.makedirs(output_dir, exist_ok=True)
    y_pred = pipeline.predict(X_test)
    labels = np.unique(y_test)

    acc = accuracy_score(y_test, y_pred)
    prec = precision_score(y_test, y_pred, average="weighted", zero_division=0)
    rec = recall_score(y_test, y_pred, average="weighted", zero_division=0)
    f1 = f1_score(y_test, y_pred, average="weighted", zero_division=0)

    cm = confusion_matrix(y_test, y_pred, labels=labels)
    plt.figure(figsize=(6, 5))
    sns.heatmap(cm, annot=True, fmt="d", cmap="Blues", xticklabels=labels, yticklabels=labels)
    plt.title(f"Confusion Matrix - {model_name}", fontsize=12)
    plt.xlabel("Nhãn dự đoán (Predicted)")
    plt.ylabel("Nhãn thực tế (True)")
    plt.tight_layout()

    cm_file = os.path.join(output_dir, f"cm_{model_name.replace(' ', '_').replace('+', 'x')}.png")
    plt.savefig(cm_file, dpi=300)
    plt.close()

    metrics = {
        "Mô hình": model_name,
        "Accuracy": acc,
        "Precision": prec,
        "Recall": rec,
        "F1-Score": f1
    }

    report = classification_report(y_test, y_pred, target_names=labels, zero_division=0)
    return metrics, report

def compare_and_save_results(metrics_list, output_dir="evaluation_results"):
    os.makedirs(output_dir, exist_ok=True)
    df_comparison = pd.DataFrame(metrics_list)

    excel_path = os.path.join(output_dir, "model_comparison.xlsx")
    csv_path = os.path.join(output_dir, "model_comparison.csv")
    df_comparison.to_excel(excel_path, index=False)
    df_comparison.to_csv(csv_path, index=False, encoding="utf-8-sig")

    df_plot = df_comparison.set_index("Mô hình")[["Accuracy", "Precision", "Recall", "F1-Score"]]
    ax = df_plot.plot(kind="bar", figsize=(12, 6), colormap="viridis", width=0.7)
    plt.title("So sánh hiệu suất giữa các phương pháp đặc trưng và thuật toán", fontsize=13)
    plt.ylabel("Điểm số")
    plt.ylim(0, 1.1)
    plt.xticks(rotation=15, ha="right")
    plt.grid(axis="y", linestyle="--", alpha=0.7)
    plt.legend(bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.tight_layout()

    chart_path = os.path.join(output_dir, "model_comparison_chart.png")
    plt.savefig(chart_path, dpi=300)
    plt.close()

    print(f"\nĐã lưu bảng kết quả so sánh tại: {excel_path}")
    print(f"Đã lưu biểu đồ so sánh tại: {chart_path}")
    return df_comparison


# HÀM CHẠY HUẤN LUYỆN VÀ LƯU MÔ HÌNH
def train_and_evaluate_all():
    X_train, X_test, y_train, y_test = load_and_preprocess_data("comments_merged.xlsx")
    pipelines = build_all_pipelines()
    metrics_list = []

    print("BẮT ĐẦU HUẤN LUYỆN VÀ ĐÁNH GIÁ CÁC MÔ HÌNH")

    saved_models_dir = os.path.join("evaluation_results", "saved_models")
    os.makedirs(saved_models_dir, exist_ok=True)

    for idx, (name, pipe) in enumerate(pipelines.items(), 1):
        print(f"\n[{idx}/5] Đang huấn luyện: {name} ...")
        pipe.fit(X_train, y_train)

        metrics, report = evaluate_single_model(name, pipe, X_test, y_test)
        metrics_list.append(metrics)

        print(f"BÁO CÁO: {name} ")
        print(report)
        print(f"Accuracy: {metrics['Accuracy']:.4f} | F1-Score: {metrics['F1-Score']:.4f}")

        saved_file = os.path.join(saved_models_dir, f"{name.replace(' ', '_').replace('+', 'x')}.pkl")
        joblib.dump(pipe, saved_file)
        print(f"    -> Đã lưu mô hình: {saved_file}")

    df_comparison = compare_and_save_results(metrics_list)
    print("BẢNG TỔNG HỢP HIỆU SUẤT:")
    print(df_comparison.to_string(index=False))

if __name__ == "__main__":
    train_and_evaluate_all()