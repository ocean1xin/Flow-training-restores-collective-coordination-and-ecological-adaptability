import os
import pandas as pd
import numpy as np
'''
source_dir =
target_dir = 
'''


if not os.path.exists(target_dir):
    os.makedirs(target_dir)

def read_fish_direction_data(file_path):
    try:
        df = pd.read_excel(file_path)
        df['dx'] = df['终点X'] - df['起点X']
        df['dy'] = df['终点Y'] - df['起点Y']
        return df
    except Exception as e:
        print(f"无法读取 {file_path}: {e}")
        return None

def calculate_ring_angle_and_category(df, pool_center):
    df['位置向量X'] = df['终点X'] - pool_center[0]
    df['位置向量Y'] = df['终点Y'] - pool_center[1]
    df['切线方向X'] = df['位置向量Y']
    df['切线方向Y'] = -df['位置向量X']
    tangent_norms = np.sqrt(df['切线方向X'] ** 2 + df['切线方向Y'] ** 2)
    df['切线方向X'] /= tangent_norms
    df['切线方向Y'] /= tangent_norms
    motion_norms = np.sqrt(df['dx'] ** 2 + df['dy'] ** 2)
    df['运动方向X'] = df['dx'] / motion_norms
    df['运动方向Y'] = df['dy'] / motion_norms
    df['环形夹角'] = np.arccos(df['切线方向X'] * df['运动方向X'] + df['切线方向Y'] * df['运动方向Y'])
    df['sign'] = (df['运动方向X'] * df['位置向量X'] + df['运动方向Y'] * df['位置向量Y'])
    df['环形夹角'] = np.where(df['sign'] < 0, df['环形夹角'], -df['环形夹角'])
    df['环形夹角_deg'] = np.degrees(df['环形夹角'])
    conditions = [
        (df['sign'] < 0) & (df['切线方向X'] * df['运动方向X'] + df['切线方向Y'] * df['运动方向Y'] > 0),
        (df['sign'] > 0) & (df['切线方向X'] * df['运动方向X'] + df['切线方向Y'] * df['运动方向Y'] > 0),
        (df['sign'] < 0) & (df['切线方向X'] * df['运动方向X'] + df['切线方向Y'] * df['运动方向Y'] < 0),
        (df['sign'] > 0) & (df['切线方向X'] * df['运动方向X'] + df['切线方向Y'] * df['运动方向Y'] < 0),
    ]
    choices = ['a', 'b', 'c', 'd']
    df['圆环类别'] = np.select(conditions, choices, default='未知')
    return df

def calculate_math_angle_and_category(df):
    df['数学坐标系角度'] = np.arctan2(df['dy'], df['dx'])
    df['数学坐标系角度_deg'] = np.degrees(df['数学坐标系角度'])
    conditions = [
        (df['数学坐标系角度_deg'] >= 0) & (df['数学坐标系角度_deg'] < 90),
        (df['数学坐标系角度_deg'] >= 90) & (df['数学坐标系角度_deg'] < 180),
        (df['数学坐标系角度_deg'] >= -180) & (df['数学坐标系角度_deg'] < -90),
        (df['数学坐标系角度_deg'] >= -90) & (df['数学坐标系角度_deg'] < 0),
    ]
    choices = [1, 2, 3, 4]
    df['数学坐标系类别'] = np.select(conditions, choices, default=0)
    return df

def process_excel_files():
    for dirpath, _, filenames in os.walk(source_dir):
        for filename in filenames:
            if filename.endswith('.xlsx'):
                file_path = os.path.join(dirpath, filename)
                print(f"处理文件: {file_path}")
                df = read_fish_direction_data(file_path)
                if df is not None:
                    pool_center = (df['终点X'].mean(), df['终点Y'].mean())
                    df = calculate_ring_angle_and_category(df, pool_center)
                    df = calculate_math_angle_and_category(df)
                    output_file_path = os.path.join(target_dir, filename)
                    df.to_excel(output_file_path, index=False)
                    print(f"结果已写入: {output_file_path}")

if __name__ == "__main__":
    process_excel_files()