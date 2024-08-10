# the y-axis is where the path begins and the x-axis is where the path ends.
import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd
import xlrd
import numpy as np  # Import numpy for NaN handling

damping_factor = 0.5

def maxelements(seq):
    max_indices = []
    if seq:
        max_val = seq[0]
        for i, val in ((i, val) for i, val in enumerate(seq) if val >= max_val):
            if val == max_val:
                max_indices.append(i)
            else:
                max_val = val
                max_indices = [i]

    return max_indices

file = "updated_original.xlsx"

G = nx.DiGraph()
_book = xlrd.open_workbook(file)
_sheet = _book.sheet_by_index(0)

for row in range(_sheet.nrows):
    if row > 0:
        _data = _sheet.row_slice(row)
        _Node1 = _data[0].value
        _Node2 = _data[1].value
        G.add_edge(_Node1, _Node2, weight=float(_data[4].value))

nodes_list = list(G.nodes())
pr = nx.pagerank_numpy(G, alpha=damping_factor)
pr_keys = list(pr.keys())
pr_values = list(pr.values())

gs = sum(pr_values)
if gs != 1:
    print("Global rank does not sum up to 1: " + str(gs))

_total_combinations = 0
_page_rank_global = {}
_page_rank_local = {}

for _node1 in nodes_list:
    _page_rank_global[_node1] = {node: np.nan for node in nodes_list}
    _page_rank_local[_node1] = {node: np.nan for node in nodes_list}

for _node1 in nodes_list:
    for _node2 in nodes_list:
        if _node1 != _node2:
            if nx.has_path(G, _node1, _node2):
                _total_combinations += 1
                print(f"############# {_total_combinations}. {_node1} --> {_node2} #############\n")
                _shortest = 99999999
                _longest = 0
                _mostInfluencedKey = " "
                _mostInfluencedValue = -1
                _all_paths = nx.all_simple_paths(G, _node1, _node2)

                _count = 0
                _all_paths_nodes = set()
                for path in _all_paths:
                    _all_paths_nodes = _all_paths_nodes.union(set(path))

                    for _item in path:
                        _item_index = pr_keys.index(_item)
                        _item_pr_value = pr_values[_item_index]

                        if _mostInfluencedValue < _item_pr_value:
                            _mostInfluencedValue = _item_pr_value
                            _mostInfluencedKey = _item

                    _count += 1
                    _path_len = len(path)
                    if _shortest > _path_len:
                        _shortest = _path_len
                    if _longest < _path_len:
                        _longest = _path_len

                if len(_all_paths_nodes) != 0:
                    G_sub = G.subgraph(list(_all_paths_nodes))
                    pr_sub = nx.pagerank_numpy(G_sub, alpha=damping_factor)
                    pr_sub_values = list(pr_sub.values())
                    pr_sub_keys = list(pr_sub.keys())
                    pr_sub_max_index = pr_sub_values.index(max(pr_sub_values))

                    _most_influential_key_and_value = f"{pr_sub_keys[pr_sub_max_index]} ({pr_sub_values[pr_sub_max_index]:.4f})"
                    _page_rank_local[_node1][_node2] = _most_influential_key_and_value

                    _most_influential_key_and_value_global = f"{_mostInfluencedKey} ({_mostInfluencedValue:.4f})"
                    _page_rank_global[_node1][_node2] = _most_influential_key_and_value_global
                else:
                    print("Most influential key and value: N/A")

                print(f"Total paths: {_count}")
                if _count != 0:
                    print(f"\tShortest path length: {_shortest}")
                    print(f"\tLongest path length: {_longest}")
                else:
                    print("\tShortest path length: N/A")
                    print("\tLongest path length: N/A")

df_local = pd.DataFrame.from_dict(_page_rank_local).sort_index(axis=0).sort_index(axis=1)
df_local.to_excel("Local.xlsx", engine='xlsxwriter')

df_global = pd.DataFrame.from_dict(_page_rank_global).sort_index(axis=0).sort_index(axis=1)
df_global.to_excel("Global.xlsx", engine='xlsxwriter')

print("Total combinations: " + str(_total_combinations))
