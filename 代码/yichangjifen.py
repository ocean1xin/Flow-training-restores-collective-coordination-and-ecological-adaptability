
import numpy as np
import pandas as pd
import os
import datetime
from sklearn.neighbors import LocalOutlierFactor
from matplotlib.colors import LinearSegmentedColormap

class MultiCriteriaOutlierDetector:


    def __init__(self, df, pool_center, pool_radius):
        self.df = df.copy()
        self.pool_center = pool_center
        self.pool_radius = pool_radius
        self.max_score = 4

        self.colors = ['#0066CC', '#3399FF', '#66CCFF', '#FFCC66', '#FF6633', '#CC0000']
        self.base_cmap = LinearSegmentedColormap.from_list("normal_to_outlier", self.colors, N=256)
        self.data_vmin = None
        self.data_vmax = None

    def calculate_voting_scores(self):
        self.df['异常投票得分'] = 0

        # 指标1: 方向类型异常 (逆时针 = 异常)
        if '方向类型' in self.df.columns:
            direction_outliers = self.df['方向类型'] == '逆时针'
            self.df['方向异常'] = direction_outliers
            self.df['异常投票得分'] += direction_outliers.astype(int)
        # 指标2: 符合要求异常 (FALSE = 异常)
        if '符合要求' in self.df.columns:
            requirement_outliers = self.df['符合要求'] == False
            self.df['要求异常'] = requirement_outliers
            self.df['异常投票得分'] += requirement_outliers.astype(int)
        # 指标3: 圆心垂线投影异常
        center_projection_outliers = self._detect_projection_outliers('center')
        self.df['圆心投影异常'] = center_projection_outliers
        self.df['异常投票得分'] += center_projection_outliers.astype(int)
        # 指标4: 群心垂线投影异常
        group_projection_outliers = self._detect_projection_outliers('group')
        self.df['群心投影异常'] = group_projection_outliers
        self.df['异常投票得分'] += group_projection_outliers.astype(int)
        # 归一化
        self.df['异常程度'] = self.df['异常投票得分'] / self.max_score
        # 颜色映射范围
        abnormality_values = self.df['异常程度'].values
        self.data_vmin = abnormality_values.min()
        self.data_vmax = abnormality_values.max()
        if self.data_vmin == self.data_vmax:
            self.data_vmin = max(0, self.data_vmin - 0.1)
            self.data_vmax = min(1, self.data_vmax + 0.1)


    def _detect_projection_outliers(self, projection_type='center', contamination=0.15):
        if projection_type == 'center':
            cols = ['起点_圆心距离一致X', '起点_圆心距离一致Y']
        else:
            cols = ['起点_群心距离一致X', '起点_群心距离一致Y']

        X_df = self.df[cols]
        valid_mask = ~X_df.isnull().any(axis=1)
        X_valid = X_df[valid_mask].values

        if len(X_valid) < 5:
            return pd.Series([False] * len(X_df), index=X_df.index)

        lof = LocalOutlierFactor(n_neighbors=min(20, max(2, len(X_valid) - 1)), contamination=contamination)
        outlier_labels = lof.fit_predict(X_valid)
        result = pd.Series([False] * len(X_df), index=X_df.index)
        result[valid_mask] = (outlier_labels == -1)
        return result

    def get_config(self):

        return {
            'pool_center_x': self.pool_center[0],
            'pool_center_y': self.pool_center[1],
            'pool_radius': self.pool_radius,
            'data_vmin': self.data_vmin,
            'data_vmax': self.data_vmax,
            'max_score': self.max_score
        }

def batch_enhanced_analysis(input_folder, output_folder=None, pool_center=None, pool_radius=300):

    if pool_center is None:
        pool_center = np.array([490.0, 233.0])
    if output_folder is None:
        output_folder = input_folder
    os.makedirs(output_folder, exist_ok=True)
    files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xlsx')]
    print(f"共检测到{len(files)}个xlsx文件，开始批量处理...")

    for fname in files:
        in_path = os.path.join(input_folder, fname)
        file_short = os.path.splitext(fname)[0]
        out_name = f"enhanced_fish_analysis_{file_short}.xlsx"
        out_path = os.path.join(output_folder, out_name)
        try:
            df = pd.read_excel(in_path)
            detector = MultiCriteriaOutlierDetector(df, pool_center, pool_radius)
            detector.calculate_voting_scores()
            config = detector.get_config()
            # 写入两个sheet
            with pd.ExcelWriter(out_path) as writer:
                detector.df.to_excel(writer, sheet_name='data', index=False)
                pd.DataFrame([config]).to_excel(writer, sheet_name='config', index=False)
            print(f"[√] 已处理: {fname} -> {out_name}")
        except Exception as e:
            print(f"[×] 处理失败: {fname}，原因: {e}")

if __name__ == '__main__':

    input_folder =
    output_folder =
    pool_center =
    pool_radius =
    batch_enhanced_analysis(input_folder, output_folder, pool_center, pool_radius)
