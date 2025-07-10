
# Define input and output file paths
import polars as pl
from collections import deque
from tqdm import tqdm
import csv
import re
import pandas as pd
input_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\01_Diagrams\Diagrams.csv'
input_inst = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\01_Diagrams\Diagrams_insts.csv'
output_file = input_path.replace('.csv',"_Out2.csv")
# results = pl.DataFrame(schema={"ObjNAME":pl.Utf8, 'ObjTYPE':pl.Utf8,'ObjREFERENCE':pl.Utf8, 'ObjREFERENCEtype':pl.Utf8})
# results2 = []
# Load the graph data
df = pl.read_csv(input_path)
df_inst = pd.read_csv(input_inst).fillna("NaN")
# results = pd.DataFrame({'From':[''], 'Path':[''], 'Destination':['']})
results = pd.DataFrame()
vessels = df.filter((pl.col("ObjTYPE") == 'SCEQUI') & (pl.col("ObjNAME").str.contains('/ASBHA-'))).select("ObjNAME").unique()


for item in vessels.rows():
    current = item[0]
    sub_system = '-' + current.split('-')[2][0:2] + '[0-9]{4}'
    print(current)
    object_set = df.filter((pl.col("ObjNAME") == current))
    object_set = object_set.with_columns(pl.col("ObjREFERENCE").alias("Path"))
    queue = df.filter((pl.col("ObjREFERENCE") != current) & (pl.col("ObjNAME") != current))
    # queue = queue.filter((pl.col("ObjTYPE") != 'BRAN') & (pl.col("ObjNAME").str.contains(sub_system)))
    # print(queue)
#  & (pl.col("ObjNAME").str.contains(sub_system))
    stop = True
    for i in range(10):
        object_set = object_set.join(queue, left_on='ObjREFERENCE', right_on='ObjNAME', how='left').drop_nulls()
        # object_set = object_set.with_columns([pl.concat_str([pl.col('Path'), pl.col('ObjREFERENCE_right')], separator='-->').alias('Path')])
        remove = object_set.select('ObjREFERENCE')
        object_set = object_set.select("ObjNAME", "ObjTYPE", 'Path', "ObjREFERENCE_right", "ObjREFERENCEtype_right")
        object_set = object_set.rename({'ObjREFERENCE_right':'ObjREFERENCE', 'ObjREFERENCEtype_right':'ObjREFERENCEtype'})
        queue = queue.join(remove, left_on='ObjNAME', right_on='ObjREFERENCE', how='anti')
        if(len(object_set.filter((pl.col("ObjREFERENCEtype").is_in(['SCBRAN', 'SCNOZZ', 'SCEQUI'])) & ~(pl.col("ObjREFERENCE").str.contains(sub_system)))) > 0):
            
            temp = object_set.filter((pl.col("ObjREFERENCEtype") == 'SCBRAN') & (pl.col("ObjREFERENCE").str.contains(sub_system))).to_pandas()
            results = pd.concat([results, temp], axis=0, ignore_index=True)
            object_set = object_set.filter(pl.col("ObjREFERENCEtype") != 'SCEQUI')


results = results.drop_duplicates(subset=['ObjNAME', 'ObjREFERENCE'])
# print(results)

results = results.merge(df_inst, how='left', left_on='ObjREFERENCE', right_on='ObjREFERENCE')
results.to_csv(output_file)




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

