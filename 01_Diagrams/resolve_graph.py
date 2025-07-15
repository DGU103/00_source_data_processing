
# Define input and output file paths
input_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\01_Diagrams\Diagrams.csv'
import polars as pl
from collections import deque
from tqdm import tqdm
import csv
import re
import pandas as pd
output_file = input_path.replace('.csv',"_Out.csv")

# Load the graph data
df = pl.read_csv(input_path)
results = pd.DataFrame({'From':[], 'Path':[], 'Destination':[]})
vessels = df.filter((pl.col("ObjNAME").str.contains(r'ASBJA-T-|ASBJA-V-'))&(pl.col("ObjTYPE") == 'SCEQUI'))

for item in vessels.rows():
    current = item[0]
    object_set = df.filter((pl.col("ObjNAME") == item[0]))
    visited_refs = df.filter((pl.col("ObjREFERENCE") == item[2]))
    # ref = item[2]
    stop = True
    # start = current
    # path = ref
    refs_df = df.filter(~df['ObjREFERENCE'].is_in(object_set['ObjNAME']))
    print(visited_refs)
    break
    while stop:
        refs_df = refs_df.filter((pl.col("ObjNAME") != current[0]) & (pl.col("ObjREFERENCE") != current[2]))
        if len(refs_df.filter(pl.col("ObjNAME") == ref).rows()) > 0:
            current = refs_df.filter(pl.col("ObjNAME") == ref).rows()[0]
            ref = current[2]
            if current[3] == 'SCEQUI' and start != ref:
                results.loc[len(results)] = [start, path, ref]
        else:
            results.loc[len(results)] = [start, path, '']
            break
        path = path + '-->' + ref
results.to_csv(output_file)




# for item in vessels.rows():
#     current = item[0]
#     ref = item[2]
#     stop = True
#     start = current
#     path = ref
#     refs_df = df
#     while stop:
#         refs_df = refs_df.filter((pl.col("ObjNAME") != current[0]) & (pl.col("ObjREFERENCE") != current[2]))
#         if len(refs_df.filter(pl.col("ObjNAME") == ref).rows()) > 0:
#             current = refs_df.filter(pl.col("ObjNAME") == ref).rows()[0]
#             ref = current[2]
#             if current[3] == 'SCEQUI' and start != ref:
#                 results.loc[len(results)] = [start, path, ref]
#         else:
#             results.loc[len(results)] = [start, path, '']
#             break
#         path = path + '-->' + ref
# results.to_csv(output_file)




# results.write_csv(output_path)

# for vessel in tqdm(vessels.rows()):
    # target_df = df.filter(pl.col("ObjNAME") == vessel[0])
    # working_df = df.filter(pl.col('ObjNAME') != vessel[0])
    # working_df.write_csv(output_path)

    # for i in range(10):
    #     target_df = target_df.join(working_df, left_on="ObjNAME", right_on='ObjREFERENCE', how="left")
    #     target_df = target_df.select('ObjNAME','ObjTYPE','ObjREFERENCE_right','ObjREFERENCEtype_right')
    #     target_df = target_df.rename({'ObjREFERENCE_right':'ObjREFERENCE','ObjREFERENCEtype_right':'ObjREFERENCEtype'})
    #     target_df = target_df.filter(pl.col("ObjREFERENCE") != vessel[0])
    #     target_df.write_csv(output_path.replace('.csv', 'target1.csv'))
    #     working_df = working_df.join(target_df, left_on='ObjREFERENCE', right_on="ObjNAME", how='anti')
    #     working_df.write_csv(output_path.replace('.csv', 'working.csv'))
    #     links = target_df.filter(pl.col('ObjREFERENCEtype') == 'SCEQUI')
    #     links.write_csv(output_path.replace('.csv', 'links.csv'))

    #     results.vstack(links)
    #     target_df = target_df.filter(pl.col("ObjREFERENCEtype") != 'SCEQUI')
    #     target_df.write_csv(output_path.replace('.csv', 'target2.csv'))
    #     asd = asd.drop_nulls()
    #     print(asd)
    #     filter_crit = target_df.select('ObjREFERENCE')
    #     print(filter_crit)
    #     working_df = working_df.join(filter_crit, on='ObjREFERENCE', how='anti')
    #     print(len(working_df))
    #     print('')
    # print(asd)
# Save results to CSV

# print(f"Saved {len(results)} paths to {output_file}")

