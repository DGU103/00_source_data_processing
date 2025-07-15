
# Define input and output file paths
import polars as pl
from collections import deque
from tqdm import tqdm
import csv
import re
import pandas as pd
input_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\01_Diagrams\Diagrams.csv'
input_inst = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\01_Diagrams\Diagrams_insts.csv'
output_inst = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\01_Diagrams\Diagrams_merge.csv'
# output_file = input_path.replace('.csv',"_Out2.csv")

df = pl.read_csv(input_path)
# df_inst = pd.read_csv(input_inst, delimiter=',', index_col=False ).fillna("NaN")
df_inst = pl.read_csv(input_inst)

results = pl.DataFrame()
# results = pd.DataFrame()
vessels = df.filter((pl.col("ObjTYPE") == 'SCEQUI') & (pl.col("ObjNAME").str.contains('/ASBHA-'))).select("ObjNAME").unique()


for item in vessels.rows():
    current = item[0]
    sub_system = '-' + current.split('-')[2][0:2] + '[0-9]{4}'
    print(current)
    object_set = df.filter((pl.col("ObjNAME") == current))
    object_set = object_set.with_columns(pl.col("ObjNAME").alias("Path"))
    queue = df.filter((pl.col("ObjREFERENCE") != current) & (pl.col("ObjNAME") != current))
    stop = True
    for i in range(10):
        object_set = object_set.join(queue, left_on='ObjREFERENCE', right_on='ObjREFERENCE', how='left').drop_nulls()
        remove = object_set.select('ObjREFERENCE')
        object_set = object_set.with_columns([pl.concat_str([pl.col('Path'), pl.col('ObjREFERENCE')], separator='-->').alias('Path')])

        object_set = object_set.select("ObjNAME", "ObjTYPE", 'Path', "ObjNAME_right", "ObjTYPE_right")
        object_set = object_set.rename({'ObjNAME_right':'ObjREFERENCE', 'ObjTYPE_right':'ObjREFERENCEtype'})
        queue = queue.join(remove, left_on='ObjREFERENCE', right_on='ObjREFERENCE', how='anti')
        # print(object_set)
        if(len(object_set.filter((pl.col("ObjREFERENCEtype").is_in(['SCINST', 'SCOINS'])))) > 0):
            temp = object_set.filter(pl.col("ObjREFERENCEtype").is_in(['SCINST', 'SCOINS']))
            results = pl.concat([results, temp])
            # object_set = object_set.filter(pl.col("ObjREFERENCEtype") != 'SCEQUI')


results = results.unique(subset=['ObjNAME', 'ObjREFERENCE'])
# merge = pd.merge(results, df_inst,how='left', on='ObjREFERENCE')
print(results)
# print(results)
# print(df_inst)
# results = results.join(df_inst, how='left', on='ObjREFERENCE')
results.write_csv(output_inst)



