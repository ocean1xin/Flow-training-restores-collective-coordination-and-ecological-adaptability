import numpy as np
import pandas as pd
import os

from fishingpackage import jisuanku
from fishingpackage import quadrant_cooperation

# ===================== 配置 =====================
POOL_CENTER =
POOL_RADIUS =
INPUT_PATH

def get_output_folder(excel_path: str) -> str:
    base_name = os.path.basename(excel_path)
    file_name_without_ext = os.path.splitext(base_name)[0]
    parent_dir = os.path.dirname(excel_path)
    output_folder = os.path.join(parent_dir, file_name_without_ext)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder

def process_and_save(excel_path: str, pool_center, save_path: str = None) -> str:

    try:
        data = pd.read_excel(excel_path)
    except Exception as e:
        print(f"[×] 读取数据失败: {excel_path} -> {e}")
        return None

    df = jisuanku.batch_calculate_all(data, pool_center)
    if df is None:
        print(f"[×] batch_calculate_all 返回None，跳过: {excel_path}")
        return None
    df = jisuanku.calculate_circular_behavior(df, pool_center)
    if df is None:
        print(f"[×] calculate_circular_behavior 返回None，跳过: {excel_path}")
        return None

    if save_path is None:
        output_folder = get_output_folder(excel_path)
        base_name = os.path.basename(excel_path)
        file_name_without_ext = os.path.splitext(base_name)[0]
        save_path = os.path.join(output_folder, f"{file_name_without_ext}_calculated.xlsx")
    try:
        df.to_excel(save_path, index=False)
        print(f"[√] 已保存计算结果到: {save_path}")
        return save_path
    except Exception as e:
        print(f"[×] 保存Excel失败: {save_path} -> {e}")
        return None

def process_single_excel(file_path: str) -> dict:

    print(f"\n===== 处理文件: {file_path} =====")
    calculated_excel = process_and_save(file_path, POOL_CENTER)
    if calculated_excel is None:
        print(f"[×] 处理失败，跳过后续: {file_path}")
        return None

    output_dir = get_output_folder(file_path)
    try:
        analyzed_df_xiangxianfa = quadrant_cooperation.analyze_quadrant_cooperation(
            file_path=calculated_excel,
            pool_center=POOL_CENTER,
            output_path=output_dir
        )
        if hasattr(analyzed_df_xiangxianfa, 'to_excel'):
            stat_path = os.path.join(output_dir, "cooperation_stats.xlsx")
            analyzed_df_xiangxianfa.to_excel(stat_path, index=False)
            print(f"[√] 群体协作统计已保存: {stat_path}")
        else:
            print(f"[!] 群体协作分析未返回有效DataFrame")
    except Exception as e:
        print(f"群体协作统计失败: {e}")

    return {
        'input': file_path,
        'calculated_excel': calculated_excel,
    }

def batch_process_folder(input_dir: str) -> list:

    excel_files = []
    for name in os.listdir(input_dir):
        lower = name.lower()
        if lower.endswith(('.xlsx', '.xls')) and not name.startswith('~$') and not lower.endswith('_calculated.xlsx'):
            excel_files.append(os.path.join(input_dir, name))
    if not excel_files:
        print(f"[!] 目录内未找到可处理的 Excel：{input_dir}")
        return []
    excel_files.sort()
    print(f"[i] 将处理 {len(excel_files)} 个文件...")
    results = []
    for path in excel_files:
        try:
            res = process_single_excel(path)
            if res is not None:
                results.append(res)
        except Exception as e:
            print(f"[×] 处理失败: {path} -> {e}")
    return results

if __name__ == '__main__':
    if os.path.isdir(INPUT_PATH):
        batch_results = batch_process_folder(INPUT_PATH)
        print(f"\n=== 批量完成，共成功 {len(batch_results)} 个 ===")
    elif os.path.isfile(INPUT_PATH):
        result = process_single_excel(INPUT_PATH)
        print(f"\n=== 单文件处理完成 ===\n{result}")
    else:
        raise FileNotFoundError(f"输入路径不存在：{INPUT_PATH}")