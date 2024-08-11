import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

fontsize=5

global_file_path = 'Global.xlsx'
local_file_path = 'Local.xlsx'

global_adj_matrix = pd.read_excel(global_file_path, index_col=0)
local_adj_matrix = pd.read_excel(local_file_path, index_col=0)

def extract_value(cell):
    try:
        return float(cell.split('(')[-1].replace(')', ''))
    except:
        return np.nan


global_adj_matrix_numeric = global_adj_matrix.applymap(extract_value)
local_adj_matrix_numeric = local_adj_matrix.applymap(extract_value)

divergence_matrix = local_adj_matrix_numeric - global_adj_matrix_numeric

mask = divergence_matrix.isna()

vmin = np.nanmin(divergence_matrix.values)
vmax = np.nanmax(divergence_matrix.values)

fig, ax = plt.subplots(figsize=(10, 8))

cax = ax.imshow(divergence_matrix, cmap="tab20c", vmin=vmin, vmax=vmax)

fig.colorbar(cax)

ax.set_xlabel('Target vertices', fontsize=fontsize)
ax.set_ylabel('Source vertices', fontsize=fontsize)

ax.set_xticks(range(len(divergence_matrix.columns)))
ax.set_xticklabels(divergence_matrix.columns, rotation=90, ha='center', fontsize=fontsize)
ax.set_yticks(range(len(divergence_matrix.index)))
ax.set_yticklabels(divergence_matrix.index, rotation=0, ha='right', fontsize=fontsize)

ax.xaxis.set_ticks_position('bottom')
ax.xaxis.set_label_position('bottom')

ax.grid(False)

ax.set_xticks(np.arange(-0.5, len(divergence_matrix.columns), 1), minor=True)
ax.set_yticks(np.arange(-0.5, len(divergence_matrix.index), 1), minor=True)
ax.grid(which='minor', color='black', linestyle='-', linewidth=0.1)
ax.tick_params(which='minor', size=0)

plt.title('Divergence of global and local ranks', fontsize=10, pad=10)
plt.tight_layout()
plt.show()
