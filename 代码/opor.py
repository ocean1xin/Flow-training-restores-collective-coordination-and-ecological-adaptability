import os
import pandas as pd
import numpy as np
'''
source_file =
output_file =
'''


def calc_polarization_and_rotational_order(df):

    df['dx'] = df['终点X'] - df['起点X']
    df['dy'] = df['终点Y'] - df['起点Y']


    norm = np.sqrt(df['dx']**2 + df['dy']**2)
    df['ux'] = df['dx'] / (norm + 1e-9)
    df['uy'] = df['dy'] / (norm + 1e-9)


    cx, cy = df['起点X'].mean(), df['起点Y'].mean()


    sum_ux, sum_uy = df['ux'].sum(), df['uy'].sum()
    polarization = np.sqrt(sum_ux**2 + sum_uy**2) / len(df)


    df['rx'] = df['终点X'] - cx
    df['ry'] = df['终点Y'] - cy
    rnorm = np.sqrt(df['rx']**2 + df['ry']**2)
    df['rot_term'] = (df['rx'] * df['uy'] - df['ry'] * df['ux']) / (rnorm + 1e-9)
    rotational_order = np.abs(df['rot_term'].sum()) / len(df)


    result = pd.DataFrame({
        "指标": ["极化有序度", "旋转有序度"],
        "数值": [polarization, rotational_order]
    })
    return result

if __name__ == "__main__":
    df = pd.read_excel(source_file, usecols=['起点X','起点Y','终点X','终点Y'])
    result = calc_polarization_and_rotational_order(df)
    result.to_excel(output_file, index=False)
    print(result)
    print(f"结果已写入: {output_file}")