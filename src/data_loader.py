import pandas as pd
import joblib
from sklearn.preprocessing import StandardScaler

def load_solar_data(file_path: str, target_col: str = 'target'):
    """Excelからデータを読み込み、特徴量(X)と目的変数(y)に分割する"""
    df = pd.read_excel(file_path)
    # ノートブックのロジックに基づき不要なカラムを削除
    X = df.drop(columns=[target_col])
    y = df[target_col]
    return X, y

def apply_scaling(X, scaler_path: str = 'checkpoints/scalers.pkl', is_train: bool = False):
    """StandardScalerの適用。学習時はsave、推論時はloadする"""
    if is_train:
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)
        joblib.dump(scaler, scaler_path)
    else:
        scaler = joblib.load(scaler_path)
        X_scaled = scaler.transform(X)
    return X_scaled, scaler
