import pandas as pd
from collections import defaultdict, deque
import re
from tqdm import tqdm
def find_all_paths(input_csv, output_csv, start_object, start_type, end_object, end_type):
    # Load data
    df = pd.read_csv(input_csv)
    # Build graph: object -> list of (reference, reference_type)
    graph = defaultdict(list)
    obj_types = {}
    for _, row in df.iterrows():
        obj = row['ObjNAME']
        obj_type = row['ObjTYPE']
        ref = row['ObjREFERENCE']
        ref_type = row['ObjREFERENCEtype']
        obj_types[obj] = obj_type
        if pd.notna(ref) and ref != '':
            graph[obj].append((ref, ref_type))

    # Find all start and end nodes matching the conditions
    start_nodes = [obj for obj, typ in obj_types.items() if re.match(start_object,obj) and typ == start_type]
    end_nodes = set(obj for obj, typ in obj_types.items() if re.match(end_object,obj) and typ == end_type)

    results = []

    # BFS to find all paths, avoid cycles
    for start in tqdm(start_nodes):
        print(start)
        queue = deque()
        queue.append((start, [start]))
        while queue:
            current, path = queue.popleft()
            if current in end_nodes and len(path) > 1:
                results.append({
                    'from': path[0],
                    'to': current,
                    'path': ' -> '.join(path[1:-1]) if len(path) > 2 else ''
                })
            for neighbor, _ in graph.get(current, []):
                if neighbor not in path:  # avoid cycles
                    queue.append((neighbor, path + [neighbor]))

    # Save results
    results_df = pd.DataFrame(results)
    results_df.to_csv(output_csv, index=False)

# Example usage:
# find_all_paths('input.csv', 'output.csv', 'some string', 'some second string', 'target string', 'target type')
input_path = r'C:\Users\mch107\Downloads\output\Diagrams.csv'
output_path = r'C:\Users\mch107\Downloads\output\Out_Diagrams.csv'

find_all_paths(input_path, output_path, '.{1,}(-V-|-T-).{1,}', 'SCEQUI', '.{1,}-P-.{1,}', 'SCEQUI')