#!IMPORTANT Social Network Analysis for the most influential nodes
import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd
import xlrd
import random


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


file = "Whole.xlsx"

G = nx.DiGraph()
_book = xlrd.open_workbook(file)
_sheet = _book.sheet_by_index(1)


for row in range(_sheet.nrows):
    if row > 0:
        _data = _sheet.row_slice(row)
        _Node1 = _data[0].value
        _Node2 = _data[1].value
        G.add_edge(_Node1, _Node2, weight=float(_data[2].value))


nodes_list = list(G.nodes())
pr = nx.pagerank_numpy(G, alpha=0.65)
pr_list = list(pr)
pr_keys = list(pr.keys())
pr_values = list(pr.values())



gs = sum(pr_values)
if gs != 1:
    print("global rank does not sum up to 1: " + str(gs))

_total_combinations = 0
_page_rank_global = {}
_page_rank_local = {}

print(pr)
option = {'node_color': 'pink', 'node_size': 1600, 'width': 1, 'font_size': 5, 'edge_color': 'gray',
          'with_labels': True}

for _node1 in nodes_list:
    if _node1 not in _page_rank_global:
        _page_rank_global[_node1] = {}
        _page_rank_local[_node1] = {}

    for _node2 in nodes_list:
        if _node1 != _node2:
            if nx.has_path(G, _node1, _node2):
                _total_combinations += 1
                print(
                    "############# " + str(_total_combinations) + ". " + _node1 + " --> " + _node2 + " #############\n")
                _shortest = 99999999
                _longest = 0
                _mostInfluencedKey = " "
                _mostInfluencedValue = -1
                _all_paths = nx.all_simple_paths(G, _node1, _node2)

                _count = 0
                indices = 0
                _all_paths_nodes = set()
                for path in _all_paths:
                    _all_paths_nodes = _all_paths_nodes.union(set(path))
                    # print(path)

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

                indices = [i for i, x in enumerate(pr_values) if x == _mostInfluencedValue]
                if len(indices) > 1:
                    print("these two in global pagerank: ", indices)

                print("\nSummary:")
                if len(_all_paths_nodes) != 0:
                    G_sub = G.subgraph(list(_all_paths_nodes))
                    pr_sub = nx.pagerank_numpy(G_sub, alpha=0.65)
                    pr_sub_list = list(pr_sub)
                    pr_sub_keys = list(pr_sub.keys())
                    pr_sub_values = list(pr_sub.values())

                    ls = sum(pr_sub_values)
                    if ls != 1:
                        print(pr_sub)
                        print("Local rank does not sum up to 1: " + str(ls))

                    pr_sub_max_index = pr_sub_values.index(max(pr_sub_values))
                   # pr_sub_max_index1 = maxelements(pr_sub_values)

                    _most_influential_key_and_value = ""

                   # for j in pr_sub_max_index1:
                   #     print("with list: ", pr_sub_keys[j])
                   #     print("with list: ", pr_sub_values[j])
                   #     print("without List: ", pr_sub_keys[pr_sub_max_index])
                   #     print("without List: ", pr_sub_values[pr_sub_max_index])
                   #     if pr_sub_keys[j]!=pr_sub_keys[pr_sub_max_index]:
                   #         int("not equal...")

                    _most_influential_key_and_value += pr_sub_keys[pr_sub_max_index] + " (" + "%.4f" % \
                                                           pr_sub_values[
                                                               pr_sub_max_index] + ")"

                    _page_rank_local[_node1][_node2] = _most_influential_key_and_value
                    print("Most influential key and value (calculated locally): " + _most_influential_key_and_value)

                    _most_influential_key_and_value_global = _mostInfluencedKey + " (%.4f" % _mostInfluencedValue + ")"

                    _page_rank_global[_node1][_node2] = _most_influential_key_and_value_global

                    print(
                        "Most influential key and value (calculated globally): " + _most_influential_key_and_value_global)
                else:
                    print("Most influential key and value: N/A")

                print("Total paths: " + str(_count))
                if _count != 0:
                    print("\tShortest path length: " + str(_shortest))
                    print("\tLongest path length: " + str(_longest))
                else:
                    print("\tShortest path length: N/A")
                    print("\tLongest path length: N/A")



df_local = pd.DataFrame.from_dict(_page_rank_local).sort_index(axis=0).sort_index(axis=1)
df_local.to_excel("Local.xlsx", engine='xlsxwriter')

df_global = pd.DataFrame.from_dict(_page_rank_global).sort_index(axis=0).sort_index(axis=1)
df_global.to_excel("Global.xlsx", engine='xlsxwriter')

print("Total combinations: " + str(_total_combinations))

