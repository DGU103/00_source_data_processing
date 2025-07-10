import pandas as pd
from collections import defaultdict, deque
import csv

def find_paths(input_csv, output_csv, start_object, start_type, end_object, end_type):
    # Step 1: Load data and build graph
    df = pd.read_csv(input_csv).fillna('')  # Fill NaNs with empty strings
    graph = defaultdict(list)

    # Build graph: node = (object, object_type), edges to (reference, reference_type)
    for _, row in df.iterrows():
        src = (row['ObjNAME'], row['ObjTYPE'])
        tgt = (row['ObjREFERENCE'], row['ObjREFERENCEtype'])
        if tgt != ('', ''):
            graph[src].append(tgt)

    # Step 2: DFS to find all paths from start to end
    start_node = (start_object, start_type)
    end_node = (end_object, end_type)
    results = []

    def dfs(current, path, visited):
        if current == end_node:
            results.append({
                'from': start_node,
                'to': end_node,
                'path': [node for node in path[1:-1]]  # exclude from/to
            })
            return
        for neighbor in graph.get(current, []):
            if neighbor not in visited:
                visited.add(neighbor)
                dfs(neighbor, path + [neighbor], visited)
                visited.remove(neighbor)

    if start_node in graph:
        dfs(start_node, [start_node], {start_node})

    # Step 3: Save paths to CSV
    with open(output_csv, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['from', 'to', 'path'])
        writer.writeheader()
        for row in results:
            writer.writerow({
                'from': f'{row["from"][0]}|{row["from"][1]}',
                'to': f'{row["to"][0]}|{row["to"][1]}',
                'path': ' -> '.join([f'{x[0]}|{x[1]}' for x in row['path']])
            })

    print(f'Done! Found {len(results)} paths.')

# Example usage:
input_path = r'C:\Users\mch107\Downloads\output\Diagrams.csv'
output_path = r'C:\Users\mch107\Downloads\output\Out_Diagrams.csv'

find_paths(input_path, output_path, 'ObjNAME', 'SCEQUI', 'ObjNAME', 'SCEQUI')