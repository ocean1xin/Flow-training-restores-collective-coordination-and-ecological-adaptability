import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt




folder_path =


pattern = r'details(\d+bl).*_frame_(\d+)\.xlsx'


def plot_rose_diagram_on_ax(angles_deg, ax, frameid, mean_value):
    angles_rad = np.radians(angles_deg.dropna())
    n_bins = 36  # 玫瑰图分36份，每10度一份
    hist, bin_edges = np.histogram(angles_rad, bins=n_bins, range=(0, 2*np.pi))
    width = 2 * np.pi / n_bins
    ax.bar(bin_edges[:-1], hist, width=width, bottom=0.0, color='steelblue', alpha=0.7)
    ax.set_theta_zero_location('N')
    ax.set_theta_direction(-1)
    ax.set_yticklabels([])
    ax.text(0.5, -0.37, f'Frame: {frameid}', transform=ax.transAxes,
            fontsize=15, ha='center', va='top', color='black', fontweight='bold')
    ax.text(0.5, -0.50, f'Mean: {mean_value:.1f}°', transform=ax.transAxes,
            fontsize=15, ha='center', va='top', color='black', fontweight='bold')


file_groups = {}
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and filename.startswith('details'):
        match = re.search(pattern, filename)
        if match:
            group = match.group(1)  # 组名 eg 2bl
            frame = match.group(2)  # 帧号
            file_path = os.path.join(folder_path, filename)
            if group not in file_groups:
                file_groups[group] = []
            file_groups[group].append((file_path, frame))

for group in file_groups:
    file_groups[group].sort(key=lambda x: int(x[1]))  # 帧号排序

for group, filelist in file_groups.items():
    num_files = len(filelist)
    if num_files == 0:
        continue

    ncols = 4
    nrows = (num_files + ncols - 1) // ncols

    figsize_unit = 15
    fig_height = max(nrows * figsize_unit, 10)
    fig_width = ncols * figsize_unit

    fig, axes = plt.subplots(nrows, ncols, subplot_kw=dict(projection='polar'), figsize=(fig_width, fig_height))
    axes = np.atleast_1d(axes).flatten()

    for idx, (fpath, frameid) in enumerate(filelist):
        ax = axes[idx]
        try:
            df = pd.read_excel(fpath)
            if '环形夹角_deg' not in df.columns:
                print(f'File {fpath} missing "环形夹角_deg", skipped')
                continue
            angles = df['环形夹角_deg']
            mean_val = angles.mean()
            plot_rose_diagram_on_ax(angles, ax, frameid, mean_val)
        except Exception as e:
            print(f'Error in {fpath}: {e}')
            continue


    for i in range(num_files, nrows * ncols):
        fig.delaxes(axes[i])

    fig.suptitle(f"{group} group direction angle rose diagrams", fontsize=22, fontweight='bold')
    plt.tight_layout(rect=[0, 0.18, 1, 0.93], h_pad=2.5, w_pad=0.3)
    base_name = os.path.join(folder_path, f'{group}_rose_diagrams')
    fig.savefig(base_name + ".png", dpi=200)
    fig.savefig(base_name + ".pdf")
    fig.savefig(base_name + ".svg")
    plt.close(fig)

print("结束")
