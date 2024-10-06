#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Quarterly Conformance

### Import required Libraries


```python
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from collections import ChainMap
```

### Set Working Path


```python
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs'
os.chdir(folder_path)
```

### Import Cluster BH Counters


```python
df = sorted(glob('*.zip'))
# concat all the Cluster DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),encoding='unicode_escape',\
        skipfooter=1,engine='python',\
        usecols=['Date','Time','GCell Group',\
        '_CallSetup TCH GOS(%)_D','_CallSetup TCH GOS(%)_N',\
        '_GOS-SDCCH(%)_D','_GOS-SDCCH(%)_N',\
        '_Mobility TCH GOS(%)_D','_Mobility TCH GOS(%)_N',\
        '_DCR_D','_DCR_N',\
        '_RxQual Index DL_1','_RxQual Index DL_2',\
        '_RxQual Index UL_1','_RxQual Index UL_2',\
        'N_OutHSR_D','N_OutHSR_N',\
        'CSSR_Non Blocking_1_N','CSSR_Non Blocking_1_D',\
        'CSSR_Non Blocking_2_N','CSSR_Non Blocking_2_D'],\
        parse_dates=["Date"]) for file in df)).\
        sort_values('Date')
```

### Filter only Clusters


```python
# Filter only cluster , remove city, region, sub region
df1=df0[df0['GCell Group']\
        .str.contains('|'.join(['_Rural','_Urban','RURAL','_URBAN']))].copy().reset_index(drop=True)
```

### Sub Region Defination


```python
df2 = ChainMap(dict.fromkeys(['GUJRANWALA_CLUSTER_01_Rural',
                            'GUJRANWALA_CLUSTER_01_Urban',
                            'GUJRANWALA_CLUSTER_02_Rural','GUJRANWALA_CLUSTER_02_Urban',
                            'GUJRANWALA_CLUSTER_03_Rural','GUJRANWALA_CLUSTER_03_Urban',
                            'GUJRANWALA_CLUSTER_04_Rural','GUJRANWALA_CLUSTER_04_Urban',
                            'GUJRANWALA_CLUSTER_05_Rural','GUJRANWALA_CLUSTER_05_Urban',
                            'GUJRANWALA_CLUSTER_06_Rural','KASUR_CLUSTER_01_Rural',
                            'KASUR_CLUSTER_02_Rural','KASUR_CLUSTER_03_Rural',
                            'KASUR_CLUSTER_03_Urban','LAHORE_CLUSTER_01_Rural',
                            'LAHORE_CLUSTER_01_Urban',
                            'LAHORE_CLUSTER_02_Rural','LAHORE_CLUSTER_02_Urban',
                            'LAHORE_CLUSTER_03_Rural','LAHORE_CLUSTER_03_Urban',
                            'LAHORE_CLUSTER_04_Urban','LAHORE_CLUSTER_05_Rural',
                            'LAHORE_CLUSTER_05_Urban',
                            'LAHORE_CLUSTER_06_Rural','LAHORE_CLUSTER_06_Urban',
                            'LAHORE_CLUSTER_07_Rural','LAHORE_CLUSTER_07_Urban',
                            'LAHORE_CLUSTER_08_Rural','LAHORE_CLUSTER_08_Urban',
                            'LAHORE_CLUSTER_09_Rural','LAHORE_CLUSTER_09_Urban',
                            'LAHORE_CLUSTER_10_Urban','LAHORE_CLUSTER_11_Rural',
                            'LAHORE_CLUSTER_11_Urban','LAHORE_CLUSTER_12_Urban',
                            'LAHORE_CLUSTER_13_Urban','LAHORE_CLUSTER_14_Urban',
                            'SIALKOT_CLUSTER_01_Rural','SIALKOT_CLUSTER_01_Urban',
                            'SIALKOT_CLUSTER_02_Rural','SIALKOT_CLUSTER_02_Urban',
                            'SIALKOT_CLUSTER_03_Rural','SIALKOT_CLUSTER_03_Urban',
                            'SIALKOT_CLUSTER_04_Rural','SIALKOT_CLUSTER_05_Rural',
                            'SIALKOT_CLUSTER_05_Urban','SIALKOT_CLUSTER_06_Rural',
                            'SIALKOT_CLUSTER_06_Urban','SIALKOT_CLUSTER_07_Rural',
                            'SIALKOT_CLUSTER_07_Urban'], 'Center-1'), 
                            dict.fromkeys(['DG_KHAN_CLUSTER_01_Rural',
                            'DG_KHAN_CLUSTER_02_Rural','DG_KHAN_CLUSTER_02_Urban',
                            'DI_KHAN_CLUSTER_01_Rural','DI_KHAN_CLUSTER_01_Urban',
                            'DI_KHAN_CLUSTER_02_Rural',
                            'DI_KHAN_CLUSTER_02_Urban','DI_KHAN_CLUSTER_03_Rural',
                            'FAISALABAD_CLUSTER_01_Rural',
                            'FAISALABAD_CLUSTER_02_Rural','FAISALABAD_CLUSTER_03_Rural',
                            'FAISALABAD_CLUSTER_04_Rural',
                            'FAISALABAD_CLUSTER_04_Urban','FAISALABAD_CLUSTER_05_Rural',
                            'FAISALABAD_CLUSTER_05_Urban',
                            'FAISALABAD_CLUSTER_06_Rural','FAISALABAD_CLUSTER_06_Urban',
                            'JHUNG_CLUSTER_01_Rural',
                            'JHUNG_CLUSTER_01_Urban','JHUNG_CLUSTER_02_Rural',
                            'JHUNG_CLUSTER_02_Urban',
                            'JHUNG_CLUSTER_03_Rural','JHUNG_CLUSTER_03_Urban',
                           'JHUNG_CLUSTER_04_Rural',
                            'JHUNG_CLUSTER_04_Urban','JHUNG_CLUSTER_05_Rural',
                            'JHUNG_CLUSTER_05_Urban',
                            'SAHIWAL_CLUSTER_01_Rural','SAHIWAL_CLUSTER_01_Urban',
                            'SAHIWAL_CLUSTER_02_Rural',
                            'SAHIWAL_CLUSTER_02_Urban'], 'Center-2'), 
                            dict.fromkeys(['JAMPUR_CLUSTER_01_Urban','RAJANPUR_CLUSTER_01_Rural',
                           'RAJANPUR_CLUSTER_01_Urban',
                            'JAMPUR_CLUSTER_01_Rural','DG_KHAN_CLUSTER_03_Rural',
                            'DG_KHAN_CLUSTER_03_Urban',
                            'DG_KHAN_CLUSTER_04_Rural','DG_KHAN_CLUSTER_04_Urban',
                            'SAHIWAL_CLUSTER_03_Rural',
                            'SAHIWAL_CLUSTER_03_Urban','KHANPUR_CLUSTER_01_Rural',
                           'KHANPUR_CLUSTER_01_Urban',
                            'RAHIMYARKHAN_CLUSTER_01_Rural','RAHIMYARKHAN_CLUSTER_01_Urban',
                            'AHMEDPUREAST_CLUSTER_01_Rural','AHMEDPUREAST_CLUSTER_01_Urban',
                            'ALIPUR_CLUSTER_01_Rural','ALIPUR_CLUSTER_01_Urban',
                            'BAHAWALPUR_CLUSTER_01_Rural','BAHAWALPUR_CLUSTER_01_Urban',
                            'BAHAWALPUR_CLUSTER_02_Rural','SAHIWAL_CLUSTER_04_Rural',
                            'SAHIWAL_CLUSTER_04_Urban','MULTAN_CLUSTER_01_Rural',
                            'MULTAN_CLUSTER_01_Urban','MULTAN_CLUSTER_02_Rural',
                            'MULTAN_CLUSTER_02_Urban',
                            'MULTAN_CLUSTER_03_Rural','MULTAN_CLUSTER_03_Urban',
                            'RYK DESERT_Cluster_Rural',
                            'SADIQABAD_CLUSTER_01_Rural','SAHIWAL_CLUSTER_05_Rural',
                           'SAHIWAL_CLUSTER_05_Urban',
                            'SAHIWAL_CLUSTER_06_Rural','SAHIWAL_CLUSTER_06_Urban',
                            'SAHIWAL_CLUSTER_07_Rural',
                             'SAHIWAL_CLUSTER_07_Urban'], 'Center-3'),
                            dict.fromkeys(['ABBOTTABAD_CLUSTER_01_Rural','ABBOTTABAD_CLUSTER_01_Urban',
                            'AJK_CLUSTER_01_Rural','AJK_CLUSTER_01_Urban',
                            'BANNU_CLUSTER_01_Rural',
                            'CHAKWAL_CLUSTER_01_Rural','CHAKWAL_CLUSTER_01_Urban',
                            'FANA_CLUSTER_01_Rural',
                            'GTROAD_CLUSTER_01_Rural','GTROAD_CLUSTER_01_Urban',
                            'ISLAMABAD_CLUSTER_01_Rural','ISLAMABAD_CLUSTER_01_Urban',
                            'ISLAMABAD_CLUSTER_02_Rural','ISLAMABAD_CLUSTER_02_Urban',
                            'ISLAMABAD_CLUSTER_03_Rural','ISLAMABAD_CLUSTER_03_Urban',
                            'ISLAMABAD_CLUSTER_04_Rural','ISLAMABAD_CLUSTER_04_Urban',
                            'ISLAMABAD_CLUSTER_05_Rural','ISLAMABAD_CLUSTER_05_Urban',
                            'JHELUM_CLUSTER_01_Rural','JHELUM_CLUSTER_01_Urban',
                            'KOHAT_CLUSTER_01_Rural','KOHAT_CLUSTER_01_Urban','KOHAT_CLUSTER_02_Rural',
                            'MANSEHRA_CLUSTER_01_Rural',
                            'MANSEHRA_CLUSTER_01_Urban',
                            'MARDAN_CLUSTER_01_Rural','MARDAN_CLUSTER_01_Urban',
                            'MIRPUR_CLUSTER_01_Rural','MIRPUR_CLUSTER_01_Urban',
                            'MOTORWAY_CLUSTER_01_Rural',
                            'MURREE_CLUSTER_01_Rural','MURREE_CLUSTER_01_Urban',
                            'MUZAFFARABAD_CLUSTER_01_Rural','MUZAFFARABAD_CLUSTER_01_Urban',
                            'NOWSHEHRA_CLUSTER_01_Rural','NOWSHEHRA_CLUSTER_01_Urban',
                            'PESHAWAR_CLUSTER_01_Rural','PESHAWAR_CLUSTER_01_Urban',
                            'PESHAWAR_CLUSTER_02_Rural','PESHAWAR_CLUSTER_02_Urban',
                            'PESHAWAR_CLUSTER_03_Rural','PESHAWAR_CLUSTER_03_Urban',
                            'PESHAWAR_CLUSTER_04_Rural','PESHAWAR_CLUSTER_04_Urban',
                            'RAWALPINDI_CLUSTER_01_Rural','RAWALPINDI_CLUSTER_01_Urban',
                            'RAWALPINDI_CLUSTER_02_Rural','RAWALPINDI_CLUSTER_02_Urban',
                            'RAWALPINDI_CLUSTER_03_Urban','RAWALPINDI_CLUSTER_04_Rural',
                            'RAWALPINDI_CLUSTER_04_Urban','RAWALPINDI_CLUSTER_05_Rural',
                            'RAWALPINDI_CLUSTER_05_Urban',
                            'SWABI_CLUSTER_01_Rural','SWABI_CLUSTER_01_Urban',
                            'SWAT_CLUSTER_01_Rural','SWAT_CLUSTER_01_Urban',
                            'SWAT_CLUSTER_02_Rural',
                            'TALAGANG_CLUSTER_01_Rural',
                            'TAXILA_CLUSTER_01_Rural','TAXILA_CLUSTER_01_Urban'], 'North'),
                            dict.fromkeys(['CHAMAN_CLUSTER_21_Rural',
                            'CHAMAN_CLUSTER_21_Urban',
                            'DADU_CLUSTER_15_Rural',
                            'DADU_CLUSTER_15_Urban',
                            'GAWADAR_CLUSTER_20_Rural',
                            'GAWADAR_CLUSTER_20_Urban',
                            'GHOTKI_CLUSTER_09_Rural',
                            'GHOTKI_CLUSTER_09_Urban',
                            'HYDERABAD_CLUSTER_01_RURAL',
                            'HYDERABAD_CLUSTER_01_URBAN',
                            'HYDERABAD_CLUSTER_02_RURAL',
                            'HYDERABAD_CLUSTER_02_URBAN',
                            'HYDERABAD_CLUSTER_03_RURAL',
                            'HYDERABAD_CLUSTER_03_URBAN',
                            'HYDERABAD_CLUSTER_04_RURAL',
                            'HYDERABAD_CLUSTER_04_URBAN',
                            'JACOBABAD_CLUSTER_12_Rural',
                            'JACOBABAD_CLUSTER_12_Urban',
                            'KARACHI_CLUSTER_01_RURAL',
                            'KARACHI_CLUSTER_01_URBAN',
                            'KARACHI_CLUSTER_02_RURAL',
                            'KARACHI_CLUSTER_02_URBAN',
                            'KARACHI_CLUSTER_03_URBAN',
                            'KARACHI_CLUSTER_04_URBAN',
                            'KARACHI_CLUSTER_05_RURAL',
                            'KARACHI_CLUSTER_05_URBAN',
                            'KARACHI_CLUSTER_06_URBAN',
                            'KARACHI_CLUSTER_07_RURAL',
                            'KARACHI_CLUSTER_07_URBAN',
                            'KARACHI_CLUSTER_08_URBAN',
                            'KARACHI_CLUSTER_09_URBAN',
                            'KARACHI_CLUSTER_10_RURAL',
                            'KARACHI_CLUSTER_10_URBAN',
                            'KARACHI_CLUSTER_11_URBAN',
                            'KARACHI_CLUSTER_12_URBAN',
                            'KARACHI_CLUSTER_13_URBAN',
                            'KARACHI_CLUSTER_14_RURAL',
                            'KARACHI_CLUSTER_14_URBAN',
                            'KARACHI_CLUSTER_15_URBAN',
                            'KARACHI_CLUSTER_16_URBAN',
                            'KARACHI_CLUSTER_17_URBAN',
                            'KARACHI_CLUSTER_18_URBAN',
                            'KARACHI_CLUSTER_19_URBAN',
                            'KARACHI_CLUSTER_20_URBAN',
                            'KARACHI_CLUSTER_21_RURAL',
                            'KARACHI_CLUSTER_21_URBAN',
                            'KARACHI_CLUSTER_22_RURAL',
                            'KARACHI_CLUSTER_22_URBAN',
                            'KHUZDAR_CLUSTER_17_Rural',
                            'KHUZDAR_CLUSTER_17_Urban',
                            'LARKANA_CLUSTER_13_Rural',
                            'LARKANA_CLUSTER_13_Urban',
                            'MIRPURKHAS_CLUSTER_05_Rural',
                            'MIRPURKHAS_CLUSTER_05_Urban',
                            'MITHI_CLUSTER_04_Rural',
                            'MITHI_CLUSTER_04_Urban',
                            'MORO_CLUSTER_02_Rural',
                            'MORO_CLUSTER_02_Urban',
                            'NAWABSHAH_CLUSTER_01_Rural',
                            'NAWABSHAH_CLUSTER_01_Urban',
                            'QUETTAEAST_CLUSTER_19_Urban',
                            'QUETTAWEST_CLUSTER_18_Rural',
                            'QUETTAWEST_CLUSTER_18_Urban',
                            'SAKRAND_CLUSTER_06_Rural',
                            'SAKRAND_CLUSTER_06_Urban',
                            'SHAHDADKOT_CLUSTER_14_Rural',
                            'SHAHDADKOT_CLUSTER_14_Urban',
                            'SIBBI_CLUSTER_16_Rural',
                            'SIBBI_CLUSTER_16_Urban',
                            'SUI_CLUSTER_11_Rural',
                            'SUI_CLUSTER_11_Urban',
                            'SUKKUR_CLUSTER_10_Rural',
                            'SUKKUR_CLUSTER_10_Urban',
                            'THATTA_CLUSTER_07_Rural',
                            'THATTA_CLUSTER_07_Urban',
                            'UMERKOT_CLUSTER_03_Rural',
                            'UMERKOT_CLUSTER_03_Urban'], 'South'))

df1.loc[:, 'Region'] = df1['GCell Group'].map(df2.get)
```

### Modification Cluster Name as per requirement


```python
# Cluster is Urban or Rural, get from GCell Group
df1['Cluster Type'] = df1['GCell Group'].str.strip().str[-5:]
# Remove Urban and Rural from GCell Group
df1['Location']=df1['GCell Group'].map(lambda x: str(x)[:-6])
```

### Select Required Quater


```python
# Identify the Quater 
df1['Quater'] = pd.PeriodIndex(pd.to_datetime(df1.Date), freq='Q')
# Select Required Quater
df3=df1[df1.Quater=='2024Q4']
```


```python
df3 = df3.copy()
df3.loc[:, 'Cluster Type'] = df3['Cluster Type'].str.capitalize()
```

### Step-1 Days Sheet


```python
# add two blank columns in a dataframe
df3=df3.assign(Comments="", Step1="")
# select the required columns
df4=df3\
    [['Region','Cluster Type','Date',\
      'GCell Group','Comments','Step1']]
```

### Sum of Counters Values


```python
df5=df3.groupby(['Region','Location','Cluster Type','GCell Group','Quater'])\
    [['_CallSetup TCH GOS(%)_D','_CallSetup TCH GOS(%)_N',\
    '_GOS-SDCCH(%)_D','_GOS-SDCCH(%)_N',\
    '_Mobility TCH GOS(%)_D', '_Mobility TCH GOS(%)_N',\
      '_DCR_D','_DCR_N',\
     '_RxQual Index DL_1','_RxQual Index DL_2',\
      '_RxQual Index UL_1','_RxQual Index UL_2',\
       'N_OutHSR_D','N_OutHSR_N',\
        'CSSR_Non Blocking_1_N','CSSR_Non Blocking_1_D',\
          'CSSR_Non Blocking_2_N','CSSR_Non Blocking_2_D']]\
            .sum().reset_index()
```

### Calculate Quarter KPIs Value

### RFP KPIS Calculateion


```python
# GOS SDCCH
df5['SDCCH GoS']=(df5['_GOS-SDCCH(%)_N']/
                               df5['_GOS-SDCCH(%)_D'])*100
```


```python
#TCH GOS
df5['TCH GoS']=(df5['_CallSetup TCH GOS(%)_N']/
                            df5['_CallSetup TCH GOS(%)_D'])*100
```


```python
# MoB TCH GoS
df5['MoB GoS']=(df5['_Mobility TCH GOS(%)_N']/
                             df5['_Mobility TCH GOS(%)_D'])*100
```

### RFO KPIs Calculation


```python
# CSSR Non Blocking
df5['CSSR']=(1-(df5['CSSR_Non Blocking_1_N']/
                    df5['CSSR_Non Blocking_1_D']))*\
             (1-(df5['CSSR_Non Blocking_2_N']/
                 df5['CSSR_Non Blocking_2_D']))*100
```


```python
# DCR
df5['DCR']=(df5['_DCR_N']/
                         df5['_DCR_D'])*100
```


```python
# Handover Success Rate
df5['HSR']=(df5['N_OutHSR_N']/
                        df5['N_OutHSR_D'])*100
```


```python
# DL Quality 
df5['DL RQI']=(df5['_RxQual Index DL_1']/
                            df5['_RxQual Index DL_2'])*100
```


```python
# UL Quality
df5['UL RQI']=(df5['_RxQual Index UL_1']/
                            df5['_RxQual Index UL_2'])*100
```


```python
# Select KPIs only
df6=df5[['Region','Location',\
        'Cluster Type','CSSR','DCR','HSR','SDCCH GoS',\
           'TCH GoS','MoB GoS','DL RQI','UL RQI']]
```

### HQ Format


```python
# re-shape()
df7=df6.pivot_table(index=['Region','Location'],\
        columns='Cluster Type',\
        values=['CSSR','DCR','HSR',\
        'SDCCH GoS','TCH GoS','MoB GoS','DL RQI','UL RQI'],\
        aggfunc=sum).\
        sort_index(level=[0,1],axis=1,ascending=[True,False])
```


```python
#Rrequired sequence
df8=df7[[('CSSR', 'Urban'),('CSSR', 'Rural'),\
           ('DCR', 'Urban'),('DCR', 'Rural'),\
            ('HSR', 'Urban'),('HSR', 'Rural'),\
            ('SDCCH GoS', 'Urban'),('SDCCH GoS', 'Rural'),\
            ('TCH GoS', 'Urban'),('TCH GoS', 'Rural'),\
            ('MoB GoS', 'Urban'),('MoB GoS', 'Rural'),\
            ('DL RQI', 'Urban'),('DL RQI', 'Rural'),\
            ('UL RQI', 'Urban'),('UL RQI', 'Rural')]].reset_index()
```


    ---------------------------------------------------------------------------

    KeyError                                  Traceback (most recent call last)

    Cell In[28], line 2
          1 #Rrequired sequence
    ----> 2 df8=df7[[('CSSR', 'Urban'),('CSSR', 'Rural'),\
          3            ('DCR', 'Urban'),('DCR', 'Rural'),\
          4             ('HSR', 'Urban'),('HSR', 'Rural'),\
          5             ('SDCCH GoS', 'Urban'),('SDCCH GoS', 'Rural'),\
          6             ('TCH GoS', 'Urban'),('TCH GoS', 'Rural'),\
          7             ('MoB GoS', 'Urban'),('MoB GoS', 'Rural'),\
          8             ('DL RQI', 'Urban'),('DL RQI', 'Rural'),\
          9             ('UL RQI', 'Urban'),('UL RQI', 'Rural')]].reset_index()
         11 df8
    

    File ~\AppData\Local\anaconda3\Lib\site-packages\pandas\core\frame.py:3767, in DataFrame.__getitem__(self, key)
       3765     if is_iterator(key):
       3766         key = list(key)
    -> 3767     indexer = self.columns._get_indexer_strict(key, "columns")[1]
       3769 # take() does not accept boolean indexers
       3770 if getattr(indexer, "dtype", None) == bool:
    

    File ~\AppData\Local\anaconda3\Lib\site-packages\pandas\core\indexes\multi.py:2539, in MultiIndex._get_indexer_strict(self, key, axis_name)
       2536     self._raise_if_missing(key, indexer, axis_name)
       2537     return self[indexer], indexer
    -> 2539 return super()._get_indexer_strict(key, axis_name)
    

    File ~\AppData\Local\anaconda3\Lib\site-packages\pandas\core\indexes\base.py:5877, in Index._get_indexer_strict(self, key, axis_name)
       5874 else:
       5875     keyarr, indexer, new_indexer = self._reindex_non_unique(keyarr)
    -> 5877 self._raise_if_missing(keyarr, indexer, axis_name)
       5879 keyarr = self.take(indexer)
       5880 if isinstance(key, Index):
       5881     # GH 42790 - Preserve name from an Index
    

    File ~\AppData\Local\anaconda3\Lib\site-packages\pandas\core\indexes\multi.py:2559, in MultiIndex._raise_if_missing(self, key, indexer, axis_name)
       2557         raise KeyError(f"{keyarr} not in index")
       2558 else:
    -> 2559     return super()._raise_if_missing(key, indexer, axis_name)
    

    File ~\AppData\Local\anaconda3\Lib\site-packages\pandas\core\indexes\base.py:5938, in Index._raise_if_missing(self, key, indexer, axis_name)
       5936     if use_interval_msg:
       5937         key = list(key)
    -> 5938     raise KeyError(f"None of [{key}] are in the [{axis_name}]")
       5940 not_found = list(ensure_index(key)[missing_mask.nonzero()[0]].unique())
       5941 raise KeyError(f"{not_found} not in index")
    

    KeyError: "None of [MultiIndex([(     'CSSR', 'Urban'),\n            (     'CSSR', 'Rural'),\n            (      'DCR', 'Urban'),\n            (      'DCR', 'Rural'),\n            (      'HSR', 'Urban'),\n            (      'HSR', 'Rural'),\n            ('SDCCH GoS', 'Urban'),\n            ('SDCCH GoS', 'Rural'),\n            (  'TCH GoS', 'Urban'),\n            (  'TCH GoS', 'Rural'),\n            (  'MoB GoS', 'Urban'),\n            (  'MoB GoS', 'Rural'),\n            (   'DL RQI', 'Urban'),\n            (   'DL RQI', 'Rural'),\n            (   'UL RQI', 'Urban'),\n            (   'UL RQI', 'Rural')],\n           names=[None, 'Cluster Type'])] are in the [columns]"


### Re-Shape Data Set (Quarter KPIs)


```python
df9=pd.DataFrame(pd.melt(df6,id_vars=['Region','Location','Cluster Type'],\
                        var_name='KPI', value_name='KPI Value')).dropna()
```

### SLA Target Values


```python
df10=pd.DataFrame({
'KPI':['CSSR','CSSR','DCR','DCR','HSR','HSR','SDCCH GoS','SDCCH GoS','TCH GoS',\
       'TCH GoS','MoB GoS','MoB GoS','DL RQI','DL RQI','UL RQI','UL RQI'],
'Cluster Type':['Urban','Rural','Urban','Rural','Urban','Rural','Urban','Rural',\
                'Urban','Rural','Urban','Rural','Urban','Rural','Urban','Rural'],
'Target Value':[99.5,99,0.6,1,97.5,96,0.1,0.1,2,2,4,4,98.4,97,98.2,97.7]
})
# Transpose SLA Target (Re-Shapre)
df11 = df10.set_index(['KPI','Cluster Type']).T
```

### Merge Re-Shape Data Set & SLA Target


```python
df12= pd.merge(df9,df10\
                [['KPI','Cluster Type','Target Value']],\
                on=['KPI','Cluster Type'])
```

### Compare KPI Value(Quarter KPIs) with SLA Target


```python
df12['Comments'] = np.where(
            (((df12['KPI']=='CSSR')|\
              (df12['KPI']=='HSR') |\
              (df12['KPI']=='DL RQI') |\
              (df12['KPI']=='UL RQI'))& \
             (df12['KPI Value'] >= \
            df12['Target Value'])),'Conformance', 
     np.where(
            (((df12['KPI']=='DCR')|\
              (df12['KPI']=='SDCCH GoS')|\
              (df12['KPI']=='TCH GoS') |\
              (df12['KPI']=='MoB GoS'))& \
             (df12['KPI Value'] <= \
              df12['Target Value'])),'Conformance', 
             'Non Conformance'))
```

### Filter Non-SLA KPIs


```python
df13=df12\
    [df12.Comments=='Non Conformance'].reset_index(drop=True)
```

### Summary


```python
# pivot table
df14=df13.pivot_table(index=['KPI','Cluster Type'],\
        columns='Region',values=['KPI Value'],\
        aggfunc='count',margins=True,margins_name='ZNationwide').\
        fillna(0).apply(np.int64).iloc[:-1,:]
```


```python
# re-shape (stack)
df15=pd.DataFrame(df14.stack()).reset_index()
# re-spare (pivot_table)
df16 = df15.pivot_table(index=['Region','Cluster Type'],\
                columns='KPI', aggfunc='sum').fillna(0)

df16['Total NC KPIs']=df16.groupby(level=0, axis=1).sum()
```


```python
#sub total for each region
df17 = df16.unstack(0)
df18 = df17.columns.get_level_values('Region') != 'All'
df17.loc['subtotal'] = df17.loc[:, df18].sum()
df19 = df17.stack().swaplevel(0,1).sort_index()
```

### Export Output


```python
with pd.ExcelWriter('GSM_NPM_SLA_Q04-2024_KPIs_Summary.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    
    # SLA Target Sheet
    df11.to_excel(writer,sheet_name="SLA Target",\
                  engine='openpyxl',na_rep='N/A')
    # Summary w.r.t Tech Sub Region
    df19.to_excel(writer,sheet_name="Summary",\
                engine='openpyxl',startrow=3,startcol=4)
    
    # List of Non Conformance Clusters and KPIs
    df13.to_excel(writer,sheet_name="Non-Conformance",\
                 engine='openpyxl',na_rep='N/A',index=False)
    
    # Step-1 days filter
    df4.to_excel(writer,sheet_name="Step-1",\
                engine='openpyxl',na_rep='N/A',index=False)
    
    # Cluster BH Numerator denominator 
    df3.to_excel(writer,sheet_name='Cluster_Num_Den',\
                engine='openpyxl',na_rep='N/A',index=False)
    
    # Cluster BH Quater Level
    df8.to_excel(writer,sheet_name="2024Q4",\
                engine='openpyxl',na_rep='N/A')
```


```python
# re-set all the variable from the RAM
%reset -f
```

# Formatting 

## Import required Libraries


```python
import os
import openpyxl
from openpyxl import load_workbook
```

## Set Input File Path


```python
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs'
os.chdir(folder_path)
```

## Load Excel File


```python
# Load the workbook
wb = load_workbook('GSM_NPM_SLA_Q04-2024_KPIs_Summary.xlsx')
```

## Set Date Format (All the Sheets)


```python
from datetime import datetime
from openpyxl.styles import NamedStyle
# Define a custom date style
date_style = NamedStyle(name='custom_date_style', number_format='DD/MM/YYYY')

# Loop through each sheet in the workbook
for ws in wb:
    # Loop through each column in the sheet
    for col in ws.iter_cols():
        for cell in col:
            if isinstance(cell.value, datetime):
                cell.style = date_style
```

## SLA Target Sheet Formaating


```python
# Select the worksheet
ws = wb['SLA Target']
```

### Delete Row


```python
# Delete the row
ws.delete_rows(3)
```

### background and font color (Multi-Index Header)


```python
# set background and font color
for row in range(1,3):
    for cell in ws[row]:
        cell.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type = "solid")
        font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=11,name='Calibri Light')
        cell.font = font
```

### Apply alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

### Apply Border


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))

# Loop through each cell in the worksheet and set the border style
for row in ws.iter_rows():
    for cell in row:
        cell.border = border
```

### Apply Fix Column Width on all the Columns


```python
# Set the width of each column to 12
for column in ws.columns:
    for cell in column:
        if cell.coordinate not in ws.merged_cells:
            ws.column_dimensions[cell.column_letter].width = 14
```

### Change Tab Color


```python
# Set the tab color
ws.sheet_properties.tabColor = "0072C6"
```

## Summary Sheet Formaating


```python
# Select the worksheet
ws1 = wb['Summary']
```

### Un-merge all the cells in a sheet


```python
for merge in list(ws1.merged_cells):
    ws1.unmerge_cells(range_string=str(merge))
```

### Delete the unwanted rows


```python
# Delete the row
ws1.delete_rows(4)
ws1.delete_rows(5)
```

### Fill the Text in the Last Column 


```python
# check the last column
last_column = ws1.max_column
# Check if the last cell of row 5 is blank
if ws1.cell(row=4, column=last_column).value is None:
    # Fill the last blank cell in row 5 with 'Total NC KPIs'
    ws1.cell(row=4, column=last_column).value = 'Total NC KPIs'
```

### Fil the Cell with Text of specific row and column


```python
ws1.cell(row=4, column=5).value = 'Region'
ws1.cell(row=4, column=6).value = 'Cluster Type' 
```

### insert blank row 


```python
ws1.insert_rows(1)
```

### Merge Specific Rows and columns


```python
ws1.merge_cells(start_row=2, start_column=5, end_row=3, end_column=last_column)
```

### Fill the Merge Cell with Text


```python
# Fill the text 'CS Summary' in the merged cell
ws1.cell(row=2, column=5).value = 'CS Summary'
```

### Set the Column Width


```python
# Set the width of each column to 12
for column in ws1.columns:
    for cell in column:
        if cell.coordinate not in ws1.merged_cells:
            ws1.column_dimensions[cell.column_letter].width = 12
```

### Apply Border (user define start row and start column)


```python
# Iterate through the rows in the table
for row in list(ws1.rows)[4:]:
    # Iterate through the cells in the current row
    for cell in row[4:]:
        # Apply the border
        cell.border = border
# border already define above
```

### Apply the alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws1.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

### Set the tab color


```python
# Set the tab color
ws1.sheet_properties.tabColor = "000000"  
```

### Set Font and Background Colour For Specific Row and Column 


```python
# Access the first row starting from row 3
first_row1 = list(ws1.rows)[1]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.fill = openpyxl.styles.PatternFill(start_color="344353", end_color="344353", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=24)
    cell.font = font
```

### un-bold the Text From specific row and Column


```python
# set un-bold for all the text except row 2 and 3
from openpyxl.styles import Font
for row in ws1.iter_rows(min_row=5, max_row=ws1.max_row):
    for cell in row:
        cell.font = Font(bold=False)
```

### Set Font and Background Colour For Specific Row and Column 


```python
from openpyxl.styles import Font
# Access the first row starting from row 2
first_row = list(ws1.rows)[4]
# Iterate through the cells in the first row starting from column E
for cell in first_row[4:]:
    cell.fill = openpyxl.styles.PatternFill(start_color="acb9ca", end_color="acb9ca", fill_type = "solid")
    font = openpyxl.styles.Font(color="000000",bold=False)
    cell.font = font
```

### Change the background color From Specific Rows and Columns in a loop


```python
from openpyxl.styles import PatternFill
for row in ws1.iter_rows(min_row=6, max_row=ws1.max_row):
    if (row[0].row-6) % 3 == 0:
        color = "D9E1F2"
    elif (row[0].row-6) % 3 == 1:
        color = "BDD7EE"
    elif (row[0].row-6) % 3 == 2:
        color = "FFC000"
    else:
        color = None
    for cell in row[5:ws1.max_column-1]:
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
```

### Change the Color of Specific Column


```python
# change the color of the total NC
color = "D9E1F2" # color code for pink
for row in ws1.iter_rows(min_row=6, max_row=ws1.max_row):
    cell = row[ws1.max_column-1]
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
```

### Merge Cells in a loop


```python
for i in range(6,ws1.max_row,3):
    ws1.merge_cells(start_row=i, start_column=5, end_row=i+2, end_column=5)
```

### Wrap the Text


```python
# wrap the text of region column
for row in ws1.iter_rows(min_row=6, max_row=ws1.max_row):
    cell = row[4]
    cell.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center',vertical='center')
```

### replace Values


```python
for row in ws1.iter_rows(min_row=6, max_row=ws1.max_row):
    if row[5].value == 'subtotal':
        row[5].value = 'Subtotal'
```

### Set Height 


```python
# set height of cell
for row in ws1.iter_rows():
    ws1.row_dimensions[row[0].row].height = 14
```


```python
# Set height of a specific row
ws1.row_dimensions[4].height = 5
```

### Apply Header


```python
# apply border around 'CS Summary'
for col in ws1.iter_cols(min_col=5, max_col=ws1.max_column, min_row=2, max_row=3):
    for cell in col:
        cell.border = border
# border define above
```

### Conditional Font Setting


```python
# apply font bold every 3rd row (Apply of Subtotal Row)
from openpyxl.styles import Font
bold_font = Font(bold=True)
for row in ws1.iter_rows(min_row=8, max_row=ws1.max_row, min_col=6, max_col=ws1.max_column):
    if (row[0].row - 8) % 3 == 0:
        for cell in row:
            cell.font = bold_font
```

## Non-Conformance Sheet Formaating


```python
# Select the worksheet
ws2 = wb['Non-Conformance']
```

### Auto fix width of the columns


```python
from openpyxl.utils import get_column_letter
# Set the width of all columns to the maximum width of their contents
for column_cells in ws2.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    ws2.column_dimensions[get_column_letter(column_cells[0].column)].width = length
```

### Apply border


```python
# Loop through each cell in the worksheet and set the border style
for row in ws2.iter_rows():
    for cell in row:
        cell.border = border
# border define above
```

### Header Format


```python
for cell in ws2[1]:
    cell.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True)
    cell.font = font
```

### Apply Alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws2.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

### Set number Format


```python
for row in ws2.iter_rows(min_col=5, max_col=5):
    for cell in row:
        if isinstance(cell.value,float):
            cell.number_format = '0.00'
```

### Round Decimal Numbers


```python
'''
for row in ws2.iter_rows(min_col=5,max_col=5):
    for cell in row:
        if isinstance(cell.value,float):
            cell.value = round(cell.value,2)
'''
```

### Set Tab Colour


```python
# Set the tab color
ws2.sheet_properties.tabColor = "FF0000"  
```

## Step-1 Sheet Formaating


```python
# Select the worksheet
ws3 = wb['Step-1']
```

### Set Tab Colour


```python
# Set the tab color
ws3.sheet_properties.tabColor = "00FF00"  
```

### Auto fix width of the columns


```python
from openpyxl.utils import get_column_letter
# Set the width of all columns to the maximum width of their contents
for column_cells in ws3.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    ws3.column_dimensions[get_column_letter(column_cells[0].column)].width = length
```

### Apply border


```python
# Loop through each cell in the worksheet and set the border style
for row in ws3.iter_rows():
    for cell in row:
        cell.border = border
# border define above
```

### Header Format


```python
for cell in ws3[1]:
    cell.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True)
    cell.font = font
```

### Apply Alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws3.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

## Cluster_Num_Den Sheet Formatting


```python
# Select the worksheet
ws4 = wb['Cluster_Num_Den']
```

### Delete the Last Two Columns


```python
# Delete the last column of the sheet
ws4.delete_cols(ws4.max_column-1, 2)
```

### Set Tab Colour


```python
# Set the tab color
ws4.sheet_properties.tabColor = "FFC0CB" 
```

### Apply border


```python
# Loop through each cell in the worksheet and set the border style
for row in ws4.iter_rows():
    for cell in row:
        cell.border = border
# border define above
```

### Header Format


```python
for cell in ws4[1]:
    cell.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True)
    cell.font = font
```

### Apply Alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws4.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

### Wrap Text on the Header


```python
# wrape text of the header and center aligment
for ws4 in wb:
    for row in ws4.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center',vertical='center')
```

### Change the Font Size


```python
from openpyxl.styles import Font

font = Font(size=11)
for row in ws4.iter_rows():
    for cell in row:
        cell.font = font
```

## Set Column Width


```python
b = ['Cluster_Num_Den']

# Iterate over all sheets in the workbook
for ws in wb.worksheets:
    # Check if the sheet is in the list of selected sheets
    if ws.title in b:
        # Set the width of column B and D to 70
        for col in ['C', 'X']:
            ws.column_dimensions[col].width = 40
        # Iterate over all columns in the sheet
        for column in ws.columns:
            if column[0].column_letter not in ['C', 'X']:
                # Set the width of all other columns to 20
                ws.column_dimensions[column[0].column_letter].width = 12
```

### KPIs Sheet Formatting


```python
# Select the worksheet
ws5 = wb['2024Q4']
```

### Delete Row


```python
ws5.delete_rows(3)
```

### Set Tab Colour


```python
# Set the tab color
ws5.sheet_properties.tabColor = "A020F0" 
```

### Merge Cells


```python
# Merge cells in rows B1:B2
ws5.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
# Merge cells in rows C1:C2
ws5.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
```

### Header Format


```python
# set background and font color
for row in range(1,3):
    for cell in ws5[row]:
        cell.fill = openpyxl.styles.PatternFill(start_color="03296b", end_color="03296b", fill_type = "solid")
        font = openpyxl.styles.Font(color="FFFFFF",bold=False,size=11)
        cell.font = font
```

### Apply Border


```python
# Loop through each cell in the worksheet and set the border style
for row in ws5.iter_rows():
    for cell in row:
        cell.border = border
# border define above
```

### Set number Format


```python
for row in ws5.iter_rows(min_col=4, max_col=20):
    for cell in row:
        if isinstance(cell.value,float):
            cell.number_format = '0.00'
```

### Apply alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws5.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

### Conditional Fomatting


```python
from openpyxl.formatting.rule import CellIsRule
# Define a variable for the fill color for empty cells
empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')

# Iterate through all the rows and columns
for row in ws5.rows:
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color

# Define range_string
last_row = ws5.max_row
```


```python
range_string_D = "D3:D" + str(last_row)   # CSSR-Urban

# Define the conditional formatting rules for D column (CSSR URBAN)
ws5.conditional_formatting.add(range_string_D, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[99.3], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_D, \
                               CellIsRule(operator='between', \
                               formula=[99.3,99.5], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_D, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[99.5],\
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_E = "E3:E" + str(last_row)   # CSSR-Rural
# Define the conditional formatting rules for E column (CSSR_RURAL)
ws5.conditional_formatting.add(range_string_E, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[98.7], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_E, \
                               CellIsRule(operator='between', \
                                          formula=[98.7,99.0], \
                                          fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_E, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[99.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_F = "F3:F" + str(last_row)   # DCR-Urban

# Define the conditional formatting rules for F column (DCR-Urban)
ws5.conditional_formatting.add(range_string_F, \
                            CellIsRule(operator='lessThanOrEqual', \
                            formula=[0.6], \
                            fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_F,\
                               CellIsRule(operator='between', \
                               formula=[0.6,0.65], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_F, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.65], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

```


```python
range_string_G = "G3:G" + str(last_row)   # DCR-Rural
# Define the conditional formatting rules for G column (DCR-Rural)
ws5.conditional_formatting.add(range_string_G,\
                               CellIsRule(operator='lessThanOrEqual', \
                            formula=[1.0], \
                            fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_G, \
                            CellIsRule(operator='between', \
                            formula=[1.0,1.05], \
                            fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_G, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[1.05], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_H = "H3:H" + str(last_row)   # HSR-Urban

# Define the conditional formatting rules for H column (HSR-Urban)
ws5.conditional_formatting.add(range_string_H, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[97], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_H, \
                            CellIsRule(operator='between', \
                            formula=[97,97.5], \
                            fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_H, \
                                CellIsRule(operator='greaterThanOrEqual', \
                                formula=[97.5], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_I = "I3:I" + str(last_row)   # HSR-Rural
# Define the conditional formatting rules for I column (HSR-Rural)
ws5.conditional_formatting.add(range_string_I, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[95.6], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_I, \
                               CellIsRule(operator='between', formula=[95.6,96], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_I, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[96.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_J = "J3:J" + str(last_row)   # SDCCH GoS-Urban

# Define the conditional formatting rules for J column (SDCCH GoS-Urban)
ws5.conditional_formatting.add(range_string_J, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[0.12], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_J, \
                               CellIsRule(operator='between', \
                               formula=[0.1,0.12], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_J, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.12], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_K = "K3:K" + str(last_row)   # SDCCH GoS-Rural
# Define the conditional formatting rules for K column (SDCCH GoS-Rural)
ws5.conditional_formatting.add(range_string_K, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[0.1], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))


ws5.conditional_formatting.add(range_string_K, \
                               CellIsRule(operator='between', \
                                formula=[0.1,0.12], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))


ws5.conditional_formatting.add(range_string_K, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.12], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_L = "L3:L" + str(last_row)   # TCH GoS-Urban

# Define the conditional formatting rules for L column (TCH GoS-Urban)
ws5.conditional_formatting.add(range_string_L, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[2.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_L, \
                               CellIsRule(operator='between', \
                                formula=[2.,2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_L, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_M = "M3:M" + str(last_row)   # TCH GoS-Rural
# Define the conditional formatting rules for M column (TCH GoS-Rural)
ws5.conditional_formatting.add(range_string_M, \
                               CellIsRule(operator='lessThanOrEqual', \
                                          formula=[2.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_M, \
                               CellIsRule(operator='between', \
                                formula=[2,2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_M, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_N = "N3:N" + str(last_row)   # MoB GoS-Urban

# Define the conditional formatting rules for N column (MoB GoS-Urban)
ws5.conditional_formatting.add(range_string_N, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[4.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_N, \
                               CellIsRule(operator='between', \
                                formula=[4.0,4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_N, \
                               CellIsRule(operator='greaterThanOrEqual',\
                                formula=[4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_O = "O3:O" + str(last_row)   # MoB GoS-Rural
# Define the conditional formatting rules for O column (MoB GoS-Rural)
ws5.conditional_formatting.add(range_string_O, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[4.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws5.conditional_formatting.add(range_string_O, \
                               CellIsRule(operator='between', \
                                formula=[4.0,4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_O, \
                               CellIsRule(operator='greaterThanOrEqual',\
                                formula=[4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```


```python
range_string_P = "P3:P" + str(last_row)   # DL RQI-Urban

# Define the conditional formatting rules for P column (DL RQI-Urban)
ws5.conditional_formatting.add(range_string_P, \
                               CellIsRule(operator='lessThanOrEqual',\
                               formula=[98.0], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_P, \
                               CellIsRule(operator='between', \
                                formula=[98.0,98.4],\
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_P,\
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[98.4],\
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_Q = "Q3:Q" + str(last_row)   # DL RQI-Rural
# Define the conditional formatting rules for Q column (DL RQI-Rural)
ws5.conditional_formatting.add(range_string_Q, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[96.5], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_Q, \
                               CellIsRule(operator='between', \
                               formula=[96.5,97], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_Q,\
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[97.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_R = "R3:R" + str(last_row)   # UL RQI-Urban

# Define the conditional formatting rules for R column (UL RQI-Urban)
ws5.conditional_formatting.add(range_string_R, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[98.0], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_R, \
                               CellIsRule(operator='between', \
                               formula=[98.0,98.2], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_R, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[98.2], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```


```python
range_string_S = "S3:S" + str(last_row)   # UL RQI-Rural
# Define the conditional formatting rules for S column (UL RQI-Rural)
ws5.conditional_formatting.add(range_string_S,\
                               CellIsRule(operator='lessThanOrEqual',\
                               formula=[97.2], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws5.conditional_formatting.add(range_string_S, \
                               CellIsRule(operator='between', \
                               formula=[97.2,97.7], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws5.conditional_formatting.add(range_string_S, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[97.7], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### Hide Column


```python
# hide the 1st column
ws5.column_dimensions.group('A', hidden=True)
```

## Adjust the Column Width


```python
# Unmerge cells in column C
ws5.unmerge_cells(start_column=3, end_column=3, start_row=1, end_row=2)

# Set the width of column C to 30
ws5.column_dimensions['C'].width = 30

# Merge cells in column C again
ws5.merge_cells(start_column=3, end_column=3, start_row=1, end_row=2)
```

### Set Zoom Size


```python
for ws in wb:
    ws.sheet_view.zoomScale = 80
```

## column Width of Specific Sheets


```python
# List of sheet names to iterate over
tt = ['Non-Conformance', 'Step-1']

# Iterate over all sheets in workbook
for ws in wb.worksheets:
    # Check if the sheet is in the list tt
    if ws.title in tt:
        # Iterate over all columns in the sheet
        for column in ws.columns:
            # Get the current width of the column
            current_width = ws.column_dimensions[column[0].column_letter].width
            # Get the maximum width of the cells in the column
            length = max(len(str(cell.value)) for cell in column)+4
            # Set the width of the column to fit the maximum width, if it's greater than the current width
            if length > current_width:
                ws.column_dimensions[column[0].column_letter].width = length
```

## Insert a New Sheet (as First Sheet)


```python
ws511 =wb.create_sheet("Title Page", 0)
```

## Merge Specific Row and Columns


```python
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
```

## Fill the Merge Cells


```python
# Fill the text 'CE Utilization Report' in the merged cell
ws511.cell(row=12, column=5).value = 'GSM NPM SLA KPIs (Q.02-2024)'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=55,name='Calibri Light')
    cell.font = font
```

## Inset Image


```python
from openpyxl.drawing.image import Image

# inset the Huawei logo
img = Image('D:/Advance_Data_Sets/KPIs_Analysis/Huawei.jpg')
img.width = 7 * 15
img.height = 7 * 15
ws511.add_image(img,'E3')

# inset the PTCL logo
img1 = Image('D:/Advance_Data_Sets/KPIs_Analysis/PTCL.png')
ws511.add_image(img1,'M3')
```

## Hide the gridlines


```python
ws511.sheet_view.showGridLines = False
```

## Hide the headings


```python
ws511.sheet_view.showRowColHeaders = False
```

## Hyper Link For Title Page


```python
# loop through all sheets in the workbook and insert the hyperlink to each sheet
row = 22
for ws in wb:
    if ws.title != "Title Page":
        hyperlink_cell = ws511.cell(row=row, column=5)
        hyperlink_cell.value = ws.title
        hyperlink_cell.hyperlink = "#'{}'!A1".format(ws.title)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.border = border
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        # set the height of the cell
        ws511.row_dimensions[row].height = 15
        # set the colum width of column 5
        ws511.column_dimensions[get_column_letter(5)].width = 30
        row += 1
```

## Hyper Link For Sub Pages


```python
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.styles import borders
# Loop through all sheets in the workbook and insert the hyperlink to each sheet
for i, ws in enumerate(wb.worksheets):
    # Check if the sheet is not the Title Page
    if ws.title != "Title Page":
        # Add hyperlink to cell in the last column+2 of the sheet
        hyperlink_cell = ws.cell(row=2, column=ws.max_column+2)
        hyperlink_cell.value = "Back to Table of Contents"
        hyperlink_cell.hyperlink = "#'{}'!E{}".format("Title Page", 22)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Set border on the hyperlink column
        for row in ws.iter_rows(min_row=2, max_row=4, min_col=hyperlink_cell.column, max_col=hyperlink_cell.column):
            for cell in row:
                cell.border = border
        # Set width of the hyperlink column
        col_letter = openpyxl.utils.get_column_letter(hyperlink_cell.column)
        ws.column_dimensions[col_letter].width = 25

        # Add hyperlink to cell in the last column+2 of the sheet for next sheet
        if i < len(wb.worksheets)-1:
            next_hyperlink_cell = ws.cell(row=3, column=ws.max_column)
            next_hyperlink_cell.value = "Next Sheet"
            next_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i+1].title)
            next_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            next_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Add hyperlink to cell in the last column+2 of the sheet for previous sheet
        if i > 0:
            prev_hyperlink_cell = ws.cell(row=4, column=ws.max_column)
            prev_hyperlink_cell.value = "Previous Sheet"
            prev_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i-1].title)
            prev_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            prev_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
```

## Clear the Cell Value


```python
# Select the sheet by index
wss = wb.worksheets[1]
# Get the cell at row 4 and the last column
cell = wss.cell(row=4, column=wss.max_column)
# Clear the contents of the cell
cell.value = None
# Remove the border of the cell
cell.border = None
```

### Save the changes


```python
# Save the changes
wb.save('GSM_NPM_SLA_Q04-2024_KPIs_Summary.xlsx')
```

### Re-Set Variables


```python
re-set all the variable
%reset -f
```

## Move Final Output to Ouput Folder


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs/GSM_NPM_SLA_Q04-2024_KPIs_Summary.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
