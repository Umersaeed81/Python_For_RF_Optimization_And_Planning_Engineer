## 📦🔧 1. Importing Required Libraries


```python
import os                      # 📁 For interacting with the operating system (file paths, directory checks, etc.)
import zipfile                 # 🗜️ For handling ZIP archive files
import numpy as np             # 🔢 For numerical operations and arrays
import pandas as pd            # 🐼 For data manipulation and analysis
from pathlib import Path       # 🛤️ For object-oriented file path handling
from datetime import time      # ⏰ For working with time objects (e.g., filtering by hour)
from functools import reduce   # 🔄 For function-based reduction (e.g., combining multiple DataFrames)
from pandas import Timestamp   # 📅 For handling timestamps specifically in pandas

from pptx.util import Inches                   # 📐 For specifying size/dimensions of shapes or images in inches
from pptx import Presentation                  # 📊 For creating and editing PowerPoint presentations
from pptx.util import Pt                       # 🔠 For setting font size in points
from pptx.enum.shapes import MSO_SHAPE_TYPE    # 🧩 For identifying and handling different shape types in slides
from pptx.dml.color import RGBColor            # 🎨 For setting custom RGB colors for text/shapes
from pptx.enum.text import PP_ALIGN            # 📏 For setting paragraph alignment (left, center, right, etc.)
from pptx.chart.data import CategoryChartData  # 📈 For providing data to category charts (like bar/column charts)
```

## 2. ✅ `standard_cg_mapping`: Reusable Cell Group Standardization Dictionary


```python
# 🔄 Master CG normalization dictionary (Reusable for all tech layers)
standard_cg_mapping = {
    # 🧩 South Region – June 25 SZ (South Phase-3)
    'L9_South_June25_SZ'    : 'L9_South_June25_SZ',
    'L9_South_Jun25_SZ'     : 'L9_South_June25_SZ',
    'L9_South_Jun25_L9_SZ'  : 'L9_South_June25_SZ',
    'L9_South_Jun25_L18_SZ' : 'L9_South_June25_SZ',
    'L9_South_Jun25_L21_SZ' : 'L9_South_June25_SZ',

    # 🧩 South – Dera Allah Yar (South Phase-3)
    'L9_South_June25_DeraAllAHYAR'    : 'L9_South_June25_DeraAllAHYAR',
    'L9_South_Jun25_DeraAllAHYAR'     : 'L9_South_June25_DeraAllAHYAR',
    'L9_South_Jun25_L9_DeraAllAHYAR'  : 'L9_South_June25_DeraAllAHYAR',
    'L9_South_Jun25_L18_DeraAllAHYAR' : 'L9_South_June25_DeraAllAHYAR',
    'L9_South_Jun25_L21_DeraAllAHYAR' : 'L9_South_June25_DeraAllAHYAR',
    
    
    # 🏙️ Jacobabad (South- Phase-2)
    'L9_Jacobabad_Apr25_2G'  : 'L9_Jacobabad_Apr25_2G',
    'L9_Jacobabad_Apr25_3G'  : 'L9_Jacobabad_Apr25_2G',
    'L9_Jacobabad_Apr25_ALL' : 'L9_Jacobabad_Apr25_2G',
    'L9_Jacobabad_Apr25_L9'  : 'L9_Jacobabad_Apr25_2G',
    'L9_Jacobabad_Apr25_L18' : 'L9_Jacobabad_Apr25_2G',
    'L9_Jacobabad_Apr25_L21' : 'L9_Jacobabad_Apr25_2G',

    # 🏙️ Shikarpur (South- Phase-2)
    'L9_Control_Shikarpur_2G'  : 'L9_Control_Shikarpur_2G',
    'L9_Control_Shikarpur_3G'  : 'L9_Control_Shikarpur_2G',
    'L9_Control_Shikarpur_L18' : 'L9_Control_Shikarpur_2G',

    # 🏞️ Khuzdar (South- Phase-2)
    'L9_Khuzdar_SZ_2G'  : 'L9_Khuzdar_SZ_2G',
    'L9_Khuzdar_SZ_3G'  : 'L9_Khuzdar_SZ_2G',
    'L9_khuzdar_SZ'     : 'L9_Khuzdar_SZ_2G',
    'L9_khuzdar_L9_SZ'  : 'L9_Khuzdar_SZ_2G',
    'L9_khuzdar_L18_SZ' : 'L9_Khuzdar_SZ_2G',
    'L9_khuzdar_L21_SZ' : 'L9_Khuzdar_SZ_2G',

    # 🏙️ Major Cities – Balochistan (South- Phase-2)
    'L9_Control_MajCitiesBaluchistan_2G'     : 'L9_Control_MajCitiesBaluchistan_2G',
    'L9_Control_MajCitiesBaluchistan_3G'     : 'L9_Control_MajCitiesBaluchistan_2G',
    'L9_Control_MajCitiesBaluchistan_L18&21' : 'L9_Control_MajCitiesBaluchistan_2G',
    'L9_Control_MajCitiesBaluchistan_L18'    : 'L9_Control_MajCitiesBaluchistan_2G',
    'L9_Control_MajCitiesBaluchistan_L21'    : 'L9_Control_MajCitiesBaluchistan_2G',
    
    # ⚓ HUB (South- Phase-2)
    'L9_Hub_Apr25_2G_SZ'	:	'L9_Hub_Apr25_2G_SZ',
    'L9_Hub_Apr25_3G_SZ'	:	'L9_Hub_Apr25_2G_SZ',
    'L9_Hub_Apr25_ALL_SZ'	:	'L9_Hub_Apr25_2G_SZ',
    'L9_Hub_Apr25_L9_SZ'	:	'L9_Hub_Apr25_2G_SZ',
    'L9_Hub_Apr25_L18_SZ'	:	'L9_Hub_Apr25_2G_SZ',
    'L9_Hub_Apr25_L21_SZ'	:	'L9_Hub_Apr25_2G_SZ',
    
    
    # ⚓ HUB Control (South- Phase-2)
    'HUB_Control_2G' : 'HUB_Control_2G',
    'HUB_Control_3G' : 'HUB_Control_2G',
    'HUB_Control_4G' : 'HUB_Control_2G',
    
    # 🌴 Turbat SZ (South- Phase-2)
    'L9_Turbat_Apr25_2G_SZ'  : 'L9_Turbat_Apr25_2G_SZ',
    'L9_Turbat_Apr25_3G_SZ'  : 'L9_Turbat_Apr25_2G_SZ',
    'L9_Turbat_Apr25_ALL_SZ' : 'L9_Turbat_Apr25_2G_SZ',
    'L9_Turbat_Apr25_L9_SZ'  : 'L9_Turbat_Apr25_2G_SZ',
    'L9_Turbat_Apr25_L18_SZ' : 'L9_Turbat_Apr25_2G_SZ',
    'L9_Turbat_Apr25_L21_SZ' : 'L9_Turbat_Apr25_2G_SZ',

    # ⚓ Gwadar (South- Phase-2)
    'L2100_Gwadar_2G_City'             : 'L2100_Gwadar_2G_City',
    'L2100_Gwadar_City_3G_U2100+u900'  : 'L2100_Gwadar_2G_City',
    'L2100_Gwadar_City_4G_L1800+L2100' : 'L2100_Gwadar_2G_City',
    'GWADAR_CONTROL_L1800'             : 'L2100_Gwadar_2G_City',
    'GWADAR_CONTROL_L2100'             : 'L2100_Gwadar_2G_City',

    # 🧭 North Sunset – (North Phase-3)
    
    '2G_Sunset_NORTH_B2_SZ'		:	'2G_Sunset_NORTH_B2_SZ',
    '3G_Sunset_NORTH_B2_SZ'		:	'2G_Sunset_NORTH_B2_SZ',
    '4G_Sunset_NORTH_B2_SZ'		: 	'2G_Sunset_NORTH_B2_SZ',
    'L900_Sunset_NORTH_B2_SZ'	:	'2G_Sunset_NORTH_B2_SZ',
    'L1800_Sunset_NORTH_B2_SZ'	:	'2G_Sunset_NORTH_B2_SZ',
    'L2100_Sunset_NORTH_B2_SZ'	:	'2G_Sunset_NORTH_B2_SZ',

    # 🧭 Kohat (North Phase-3)
    '2G_Sunset_N_KOHAT_SZ'		:	'2G_Sunset_N_KOHAT_SZ',
    '3G_Sunset_N_KOHAT_SZ'		:	'2G_Sunset_N_KOHAT_SZ',
    '4G_Sunset_N_KOHAT_SZ'		:	'2G_Sunset_N_KOHAT_SZ',
    'L900_Sunset_N_KOHAT_SZ'	:	'2G_Sunset_N_KOHAT_SZ',
    'L1800_Sunset_N_KOHAT_SZ'	:	'2G_Sunset_N_KOHAT_SZ',
    'L2100_Sunset_N_KOHAT_SZ'	:	'2G_Sunset_N_KOHAT_SZ',

    # 🧭 Hangu (North Phase-3)
    '2G_Sunset_N_HANGU_SZ'		:	'2G_Sunset_N_HANGU_SZ',
    '3G_Sunset_N_HANGU_SZ'		:	'2G_Sunset_N_HANGU_SZ',
    '4G_Sunset_N_HANGU_SZ'		:	'2G_Sunset_N_HANGU_SZ',
    'L900_Sunset_N_HANGU_SZ'	:	'2G_Sunset_N_HANGU_SZ',
    'L1800_Sunset_N_HANGU_SZ'	:	'2G_Sunset_N_HANGU_SZ',
    'L2100_Sunset_N_HANGU_SZ'	:	'2G_Sunset_N_HANGU_SZ',

    # 🧭 Bannu (North Phase-3)
    '2G_Sunset_N_BANNU_SZ'		:	'2G_Sunset_N_BANNU_SZ',
    '3G_Sunset_N_BANNU_SZ'		:	'2G_Sunset_N_BANNU_SZ',
    '4G_Sunset_N_BANNU_SZ'		:	'2G_Sunset_N_BANNU_SZ',
    'L900_Sunset_N_BANNU_SZ'	:	'2G_Sunset_N_BANNU_SZ',
    'L1800_Sunset_N_BANNU_SZ'	:	'2G_Sunset_N_BANNU_SZ',
    'L2100_Sunset_N_BANNU_SZ'	:	'2G_Sunset_N_BANNU_SZ',


    
    #🧩 Center Region – June 25 SZ (Center Phase-3)

    'GSM_BZ_Center_Region_Phase-03'				        :		'GSM_BZ_Center_Region_Phase-03',
    'UMTS_BZ_Center_Region_Phase-03'			        :		'GSM_BZ_Center_Region_Phase-03',
    'LTE_BZ_Center_Region_Phase-03'				        :		'GSM_BZ_Center_Region_Phase-03',
    'LTE_BZ_Center_Region_Phase-03_PTML-PTML-L-1800'	:		'GSM_BZ_Center_Region_Phase-03',
	
	#🧩 Cluster-120 – June 25 SZ (Center Phase-3)
    'GSM_Dep_Cluster_120'					:		'GSM_Dep_Cluster_120',
    'UMTS_Dep_Cluster_120'					:		'GSM_Dep_Cluster_120',
    'LTE_Dep_Cluster_120'					:		'GSM_Dep_Cluster_120',
    'LTE_Dep_Cluster_120_PTML-L-900'		:		'GSM_Dep_Cluster_120',
    'LTE_Dep_Cluster_120_PTML-L-1800'		:		'GSM_Dep_Cluster_120',
	
	#🧩 Cluster-121 – June 25 SZ (Center Phase-3)
    'GSM_Dep_Cluster_121'					:		'GSM_Dep_Cluster_121',
    'UMTS_Dep_Cluster_121'					:		'GSM_Dep_Cluster_121',
    'LTE_Dep_Cluster_121'					:		'GSM_Dep_Cluster_121',
    'LTE_Dep_Cluster_121_PTML-L-900'		:		'GSM_Dep_Cluster_121',
    'LTE_Dep_Cluster_121_PTML-L-1800'		:		'GSM_Dep_Cluster_121',
    'LTE_Dep_Cluster_121_PTML-L-2100'		:		'GSM_Dep_Cluster_121',
	
	#🧩 Cluster-80 – June 25 SZ (Center Phase-3)
    'GSM_Dep_Cluster_80'					:		'GSM_Dep_Cluster_80',
    'UMTS_Dep_Cluster_80'					:		'GSM_Dep_Cluster_80',
    'LTE_Dep_Cluster_80'					:		'GSM_Dep_Cluster_80',
    'LTE_Dep_Cluster_80_PTML-L-900'			:		'GSM_Dep_Cluster_80',
    'LTE_Dep_Cluster_80_PTML-L-1800'		:		'GSM_Dep_Cluster_80',

    #🧩 Cluster-15 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_15'				:		'GSM_Dep_No_BZ_Cluster_15',
    'UMTS_Dep_No_BZ_Cluster_15'				:		'GSM_Dep_No_BZ_Cluster_15',
    'LTE_Dep_No_BZ_Cluster_15'				:		'GSM_Dep_No_BZ_Cluster_15',
    'LTE_Dep_No_BZ_Cluster_15_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_15',
    'LTE_Dep_No_BZ_Cluster_15_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_15',
    'LTE_Dep_No_BZ_Cluster_15_PTML-L-2100'	:		'GSM_Dep_No_BZ_Cluster_15',
	
	
	#🧩 Cluster-19 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_19'				:		'GSM_Dep_No_BZ_Cluster_19',
    'UMTS_Dep_No_BZ_Cluster_19'				:		'GSM_Dep_No_BZ_Cluster_19',
    'LTE_Dep_No_BZ_Cluster_19'				:		'GSM_Dep_No_BZ_Cluster_19',
    'LTE_Dep_No_BZ_Cluster_19_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_19',
    'LTE_Dep_No_BZ_Cluster_19_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_19',
	
 
    #🧩 Cluster-33 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_33'				:		'GSM_Dep_No_BZ_Cluster_33',
    'UMTS_Dep_No_BZ_Cluster_33'				:		'GSM_Dep_No_BZ_Cluster_33',
    'LTE_Dep_No_BZ_Cluster_33'				:		'GSM_Dep_No_BZ_Cluster_33',
    'LTE_Dep_No_BZ_Cluster_33_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_33',
    'LTE_Dep_No_BZ_Cluster_33_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_33',
    'LTE_Dep_No_BZ_Cluster_33_PTML-L-2100'	:		'GSM_Dep_No_BZ_Cluster_33',
	
	#🧩 Cluster-41 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_41'				:		'GSM_Dep_No_BZ_Cluster_41',
    'UMTS_Dep_No_BZ_Cluster_41'				:		'GSM_Dep_No_BZ_Cluster_41',
    'LTE_Dep_No_BZ_Cluster_41'				:		'GSM_Dep_No_BZ_Cluster_41',
    'LTE_Dep_No_BZ_Cluster_41_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_41',
    'LTE_Dep_No_BZ_Cluster_41_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_41',
    'LTE_Dep_No_BZ_Cluster_41_PTML-L-2100'	:		'GSM_Dep_No_BZ_Cluster_41',

    #🧩 Cluster-50 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_50'				:		'GSM_Dep_No_BZ_Cluster_50',
    'UMTS_Dep_No_BZ_Cluster_50'				:		'GSM_Dep_No_BZ_Cluster_50',
    'LTE_Dep_No_BZ_Cluster_50'				:		'GSM_Dep_No_BZ_Cluster_50',
    'LTE_Dep_No_BZ_Cluster_50_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_50',
    'LTE_Dep_No_BZ_Cluster_50_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_50',
    'LTE_Dep_No_BZ_Cluster_50_PTML-L-2100'	:		'GSM_Dep_No_BZ_Cluster_50',
	
	
	#🧩 Cluster-54 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_54'				:		'GSM_Dep_No_BZ_Cluster_54',
    'UMTS_Dep_No_BZ_Cluster_54'				:		'GSM_Dep_No_BZ_Cluster_54',
    'LTE_Dep_No_BZ_Cluster_54'				:		'GSM_Dep_No_BZ_Cluster_54',
    'LTE_Dep_No_BZ_Cluster_54_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_54',
    'LTE_Dep_No_BZ_Cluster_54_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_54',
	

    #🧩 Cluster-64 – June 25 SZ (Center Phase-3)
    'GSM_Dep_No_BZ_Cluster_64'				:		'GSM_Dep_No_BZ_Cluster_64',
    'UMTS_Dep_No_BZ_Cluster_64'				:		'GSM_Dep_No_BZ_Cluster_64',
    'LTE_Dep_No_BZ_Cluster_64'				:		'GSM_Dep_No_BZ_Cluster_64',
    'LTE_Dep_No_BZ_Cluster_64_PTML-L-900'	:		'GSM_Dep_No_BZ_Cluster_64',
    'LTE_Dep_No_BZ_Cluster_64_PTML-L-1800'	:		'GSM_Dep_No_BZ_Cluster_64',
	
	
	#🧩 SZ & BZ (Center Phase-3)
    'GSM_SZ_BZ_Center_Region_Phase-03'			        :		'GSM_SZ_BZ_Center_Region_Phase-03',
    'UMTS_SZ_BZ_Center_Region_Phase-03'			        :		'GSM_SZ_BZ_Center_Region_Phase-03',
    'LTE_SZ_BZ_Center_Region_Phase-03'			        :		'GSM_SZ_BZ_Center_Region_Phase-03',
    'LTE_SZ_BZ_Center_Region_Phase-03_PTML-PTML-L-900'	:		'GSM_SZ_BZ_Center_Region_Phase-03',
    'LTE_SZ_BZ_Center_Region_Phase-03_PTML-PTML-L-1800'	:		'GSM_SZ_BZ_Center_Region_Phase-03',
    'LTE_SZ_BZ_Center_Region_Phase-03_PTML-PTML-L-2100'	:		'GSM_SZ_BZ_Center_Region_Phase-03',
    	
    	
    #🧩 SZ (Center Phase-3)	
    'GSM_SZ_Center_Region_Phase-03'				        :		'GSM_SZ_Center_Region_Phase_03',
    'UMTS_SZ_Center_Region_Phase-03'			        :		'GSM_SZ_Center_Region_Phase_03',
    'LTE_SZ_Center_Region_Phase-03'				        :		'GSM_SZ_Center_Region_Phase_03',
    'LTE_SZ_Center_Region_Phase-03_PTML-PTML-L-900'		:		'GSM_SZ_Center_Region_Phase_03',
    'LTE_SZ_Center_Region_Phase-03_PTML-PTML-L-1800'	:		'GSM_SZ_Center_Region_Phase_03',
    'LTE_SZ_Center_Region_Phase-03_PTML-PTML-L-2100'	:		'GSM_SZ_Center_Region_Phase_03',

    
    #🧩 cluster_9 (Center Phase-4)	
    'GSM_SZ_Dep_Cluster_9'			    :		'SZ_Dep_Cluster_9',
    'UMTS_SZ_Dep_Cluster_9'			    :		'SZ_Dep_Cluster_9',
    'LTE_SZ_Dep_Cluster_9'			    :		'SZ_Dep_Cluster_9',
    'LTE_SZ_Dep_Cluster_9_PTML-L-1800'	:		'SZ_Dep_Cluster_9',
    'LTE_SZ_Dep_Cluster_9_PTML-L-2100'	:		'SZ_Dep_Cluster_9',
    'LTE_SZ_Dep_Cluster_9_PTML-L-900'	:		'SZ_Dep_Cluster_9',

    #🧩 SAMBRIAL (Center Phase-4)	
    'GSM_Sunset_SAMBRIAL'			    :		'Sunset_SAMBRIAL',
    'UMTS_Sunset_SAMBRIAL'			    :		'Sunset_SAMBRIAL',
    'LTE_Sunset_SAMBRIAL'			    :		'Sunset_SAMBRIAL',
    'LTE_Sunset_SAMBRIAL_PTML-L-1800'	:		'Sunset_SAMBRIAL',
    'LTE_Sunset_SAMBRIAL_PTML-L-2100'	:		'Sunset_SAMBRIAL',
    'LTE_Sunset_SAMBRIAL_PTML-L-900'	:		'Sunset_SAMBRIAL',

    #🧩 SIALKOT (Center Phase-4)
    'GSM_Sunset_SIALKOT'			    :		'Sunset_SIALKOT',
    'UMTS_Sunset_SIALKOT'			    :		'Sunset_SIALKOT',
    'LTE_Sunset_SIALKOT'			    :		'Sunset_SIALKOT',
    'LTE_Sunset_SIALKOT_PTML-L-1800'	:		'Sunset_SIALKOT',
    'LTE_Sunset_SIALKOT_PTML-L-2100'	:		'Sunset_SIALKOT',
    'LTE_Sunset_SIALKOT_PTML-L-900'		:		'Sunset_SIALKOT',
    
    #🧩 PESHAWAR (North Phase-4)
    '2G_SUNSET_N_PESHAWAR_SZ'		    :		'Sunset_N_PESHAWAR_SZ',
    '3G_SUNSET_N_PESHAWAR_SZ'		    :		'Sunset_N_PESHAWAR_SZ',
    '4G_SUNSET_N_PESHAWAR_SZ'		    :		'Sunset_N_PESHAWAR_SZ',
    'L900_SUNSET_N_PESHAWAR_SZ'		    :		'Sunset_N_PESHAWAR_SZ',
    'L1800_SUNSET_N_PESHAWAR_SZ'		:		'Sunset_N_PESHAWAR_SZ',
    'L2100_SUNSET_N_PESHAWAR_SZ'		:		'Sunset_N_PESHAWAR_SZ',

    #🧩 North (North Phase-4)
    '2G_SUNSET_NORTH_B3_SZ'			:		'Sunset_NORTH_B3_SZ',
    '3G_SUNSET_NORTH_B3_SZ'			:		'Sunset_NORTH_B3_SZ',
    '4G_SUNSET_NORTH_B3_SZ'			:		'Sunset_NORTH_B3_SZ',
    'L900_SUNSET_NORTH_B3_SZ'		:		'Sunset_NORTH_B3_SZ',
    'L1800_SUNSET_NORTH_B3_SZ'		:		'Sunset_NORTH_B3_SZ',
    'L2100_SUNSET_NORTH_B3_SZ'		:		'Sunset_NORTH_B3_SZ',
    
    # South SZ (South Phase-4)
    'L9_South_July25_2G_LKN_city_SZ'	:       'South_July25_2G_P4_SZ',
    'L9_South_July25_3G_LKN_city_SZ'	:       'South_July25_2G_P4_SZ',
    'L9_South_July25_4G_LKN_city_SZ'	:       'South_July25_2G_P4_SZ',
    'L9_South_July25_LKN_city_L9_SZ'	:       'South_July25_2G_P4_SZ',
    'L9_South_July25_LKN_city_L18_SZ'	:       'South_July25_2G_P4_SZ',
    'L9_South_July25_LKN_city_L21'	    :       'South_July25_2G_P4_SZ',
	
	# South Lardkana City (South Phase-4)
    'L9_South_July25_2G_Lark_city_SZ'	:  'South_July25_2G_Lark_city',
    'L9_South_July25_3G_Lark_city_SZ'	:  'South_July25_2G_Lark_city',
    'L9_South_July25_4G_Lark_city_SZ'	:  'South_July25_2G_Lark_city',
    'L9_South_July25_LARK_city_L9'	    :  'South_July25_2G_Lark_city',
    'L9_South_July25_LARK_city_L18'	    :  'South_July25_2G_Lark_city',
    'L9_South_July25_LARK_city_L21'	    :  'South_July25_2G_Lark_city',

    # Center Buffer Zone (Phase-5)
    'GSM_BZ_Center_Phase V'				    :	'BZ_Center_Phase V',
    'UMTS_BZ_Center_Phase-V'			    :	'BZ_Center_Phase V',
    'LTE_BZ_Center_Phase-V'				    :	'BZ_Center_Phase V',
    'LTE_BZ_Center_Phase-V_PTML-L-1800'		:	'BZ_Center_Phase V',
    'LTE_BZ_Center_Phase-V_PTML-L-2100'		:	'BZ_Center_Phase V',

    # Center Buffer Zone (Phase-6)
    'GSM_BZ_Center_Phase VI'			    :	'BZ_Center_Phase VI',
    'UMTS_BZ_Center_Phase-VI'			    :	'BZ_Center_Phase VI',
    'LTE_BZ_Center_Phase-VI'			    :	'BZ_Center_Phase VI',
    'LTE_BZ_Center_Phase-VI_PTML-L-1800'	:	'BZ_Center_Phase VI',
    'LTE_BZ_Center_Phase-VI_PTML-L-2100'	:	'BZ_Center_Phase VI',

    # Center Service+Buffer Zone (Phase-5)
    'GSM_SZ_BZ_Center_Phase V'			    :	'SZ_BZ_Center_Phase V',
    'UMTS_SZ_BZ_Center_Phase-V'			    :	'SZ_BZ_Center_Phase V',
    'LTE_SZ_BZ_Cente_Phase-V'			    :	'SZ_BZ_Center_Phase V',
    'LTE_SZ_BZ_Cente_Phase-V_PTML-L-1800'	:	'SZ_BZ_Center_Phase V',
    'LTE_SZ_BZ_Cente_Phase-V_PTML-L-2100'	:	'SZ_BZ_Center_Phase V',

    # Center Service+Buffer Zone (Phase-6)
    'GSM_SZ_BZ_Center_Phase VI'			        :	'SZ_BZ_Center_Phase VI',
    'UMTS_SZ_BZ_Center_Phase-VI'			    :	'SZ_BZ_Center_Phase VI',
    'LTE_SZ_BZ_Cente_Phase-VI'			        :	'SZ_BZ_Center_Phase VI',
    'LTE_SZ_BZ_Cente_Phase-VI_PTML-L-1800'		:	'SZ_BZ_Center_Phase VI',
    'LTE_SZ_BZ_Cente_Phase-VI_PTML-L-2100'		:	'SZ_BZ_Center_Phase VI',
    
    # Center Service Zonre (Phase-5) 
    'GSM_SZ_Center_Phase V'				    :	'SZ_Center_Phase V',
    'UMTS_SZ_Center_Phase-V'			    :	'SZ_Center_Phase V',
    'LTE_SZ_Center_Phase-V'				    :	'SZ_Center_Phase V',
    'LTE_SZ_Center_Phase-V_PTML-L-1800'		:	'SZ_Center_Phase V',
    'LTE_SZ_Center_Phase-V_PTML-L-2100'		:	'SZ_Center_Phase V',
    
    # Center Service Zonre (Phase-6)
    'GSM_SZ_Center_Phase VI'			       :	'SZ_Center_Phase VI',
    'UMTS_SZ_Center_Phase-VI'			       :	'SZ_Center_Phase VI',
    'LTE_SZ_Center_Phase-VI'			       :	'SZ_Center_Phase VI',
    'LTE_SZ_Center_Phase-VI_PTML-L-1800'	   :	'SZ_Center_Phase VI',
    'LTE_SZ_Center_Phase-VI_PTML-L-2100'	   :	'SZ_Center_Phase VI',

    # Chanab_Nagar_City (Phase-5)
    'GSM_Sunset_Chanab_Nagar_City'			    :	'Chanab_Nagar_City',
    'UMTS_Sunset_Chanab_Nagar_City'			    :	'Chanab_Nagar_City',
    'LTE_Sunset_Chanab_Nagar_City'			    :	'Chanab_Nagar_City',
    'LTE_Sunset_Chanab_Nagar_City_PTML-L-1800'	:	'Chanab_Nagar_City',
    'LTE_Sunset_Chanab_Nagar_City_PTML-L-2100'	:	'Chanab_Nagar_City',

    # DG_Khan_City (Phase-5)
    'GSM_Sunset_DG_KHAN_City'			    :	'DG_KHAN_City',
    'UMTS_Sunset_DG_KHAN_City'			    :	'DG_KHAN_City',
    'LTE_Sunset_DG_KHAN_City'			    :	'DG_KHAN_City',
    'LTE_Sunset_DG_KHAN_City_PTML-L-1800'	:	'DG_KHAN_City',
    'LTE_Sunset_DG_KHAN_City_PTML-L-2100'	:	'DG_KHAN_City',

    # DI_Khan_City (Phase-6)
    'GSM_Sunset_DI_KHAN_City'			    :	'DI_KHAN_City',
    'UMTS_Sunset_DI_KHAN_City'			    :	'DI_KHAN_City',
    'LTE_Sunset_DI_KHAN_City'			    :	'DI_KHAN_City',
    'LTE_Sunset_DI_KHAN_City_PTML-L-1800'	:	'DI_KHAN_City',	
    'LTE_Sunset_DI_KHAN_City_PTML-L-2100'	:	'DI_KHAN_City',

    # Fazilpur_City (Phase-5)
    'GSM_Sunset_Fazilpur_City'			    :	'Fazilpur_City',
    'UMTS_Sunset_Fazilpur_City'			    :	'Fazilpur_City',	
    'LTE_Sunset_Fazilpur_City'			    :	'Fazilpur_City',
    'LTE_Sunset_Fazilpur_City_PTML-L-1800'	:	'Fazilpur_City',
    'LTE_Sunset_Fazilpur_City_PTML-L-2100'	:	'Fazilpur_City',

    # JAMPUR_City (Phase-5)
    'GSM_Sunset_JAMPUR_City'			    :	'JAMPUR_City',
    'UMTS_Sunset_JAMPUR_City'			    :	'JAMPUR_City',
    'LTE_Sunset_JAMPUR_City'			    :	'JAMPUR_City',
    'LTE_Sunset_JAMPUR_City_PTML-L-1800'	:	'JAMPUR_City',
    'LTE_Sunset_JAMPUR_City_PTML-L-2100'	:	'JAMPUR_City',

    # MAMU_KANJAN_City (Phase-5)
    'GSM_Sunset_MAMU_KANJAN_City'			    :	'MAMU_KANJAN_City',
    'UMTS_Sunset_MAMU_KANJAN_City'			    :	'MAMU_KANJAN_City',
    'LTE_Sunset_MAMU_KANJAN_City'			    :	'MAMU_KANJAN_City',
    'LTE_Sunset_MAMU_KANJAN_City_PTML-L-1800'	:	'MAMU_KANJAN_City',
    'LTE_Sunset_MAMU_KANJAN_City_PTML-L-2100'	:    'MAMU_KANJAN_City',

    # RAJANPUR_City (Phase-5)
    'GSM_Sunset_RAJANPUR_City'			       :	'RAJANPUR_City',
    'UMTS_Sunset_RAJANPUR_City'			       :	'RAJANPUR_City',
    'LTE_Sunset_RAJANPUR_City'			       :	'RAJANPUR_City',
    'LTE_Sunset_RAJANPUR_City_PTML-L-1800'	   :	'RAJANPUR_City',
    'LTE_Sunset_RAJANPUR_City_PTML-L-2100'	   :	'RAJANPUR_City',
 
    # South Service Zone (Phase-5)
    'L9_South_July25_2G_22nd Batch_SZ'		:	'South_July25_22nd Batch_SZ',
    'L9_South_July25_3G_22nd Batch_SZ'		:	'South_July25_22nd Batch_SZ',
    'L9_South_July25_4G_22nd Batch_SZ'		:	'South_July25_22nd Batch_SZ',
    'L9_South_July25_22nd Batch_L9_SZ'		:	'South_July25_22nd Batch_SZ',
    'L9_South_July25_22nd Batch_L18_SZ'		:	'South_July25_22nd Batch_SZ',
    'L9_South_July25_22nd Batch_L21_SZ'		:	'South_July25_22nd Batch_SZ',

    # Shahdadkot (Phase-5)
    'L9_South_July25_2G_Shahdadkot_city_SZ'		:	'South_July25_Shahdadkot_city_SZ',
    'L9_South_July25_3G_Shahdadkot_city_SZ'		:	'South_July25_Shahdadkot_city_SZ',
    'L9_South_July25_4G_Shahdadkot_city_SZ'		:	'South_July25_Shahdadkot_city_SZ',
    'L9_South_July25_Shahdadkot_city_L9_SZ'		:	'South_July25_Shahdadkot_city_SZ',
    'L9_South_July25_Shahdadkot_city_L18_SZ'	:	'South_July25_Shahdadkot_city_SZ',	
    'L9_South_July25_Shahdadkot_city_L21_SZ'	:	'South_July25_Shahdadkot_city_SZ',

    # Shikarpur (Phase-5)
    'L9_South_July25_2G_Shikarpur_city_SZ'		:	'South_July25_Shikarpur_city_SZ',
    'L9_South_July25_3G_Shikarpur_city_SZ'		:	'South_July25_Shikarpur_city_SZ',
    'L9_South_July25_4G_Shikarpur_city_SZ'		:	'South_July25_Shikarpur_city_SZ',
    'L9_South_July25_Shikarpur_city_L9_SZ'		:	'South_July25_Shikarpur_city_SZ',
    'L9_South_July25_Shikarpur_city_L18_SZ'		:	'South_July25_Shikarpur_city_SZ',
    'L9_South_July25_Shikarpur_city_L21_SZ'		:	'South_July25_Shikarpur_city_SZ',
}
```

## 📂🗜️ 3. Load & Extract Data from ZIP Files


```python
def load_gsm_from_zip_folder(zip_folder, keyword, required_columns, dtype=None):
    """
    🗜️ Extracts and combines CSV files from ZIPs in a folder, filtered by keyword and required columns.

    Parameters:
        zip_folder (str or Path): 📁 Folder containing ZIP files.
        keyword (str): 🔍 Keyword to filter relevant CSV filenames (e.g., "(GSM_Cell_Hourly)").
        required_columns (list): 🧱 List of required columns to read from CSVs.
        dtype (dict): Optional 🧬 dictionary specifying data types for specific columns.

    Returns:
        pd.DataFrame: 🐼 Combined DataFrame with valid rows, or None if no match found.
    """
    zip_files = list(Path(zip_folder).glob("*.zip"))  # 📦 List all ZIP files in the folder
    dfs = []  # 📋 Initialize list to collect DataFrames

    for zip_file in zip_files:
        with zipfile.ZipFile(zip_file, 'r') as z:
            # 🎯 Filter CSV files inside the ZIP that match the keyword
            csv_files = [f for f in z.namelist() if keyword in f and f.endswith(".csv")]

            for csv_file in csv_files:
                with z.open(csv_file) as f:
                    try:
                        df = pd.read_csv(
                            f,
                            usecols=required_columns,      # 🧩 Only load required columns
                            dtype=dtype,                   # 🧬 Apply custom dtypes if given
                            parse_dates=["Date"],          # 📆 Ensure 'Date' is parsed as datetime
                            skiprows=range(6),             # 🪄 Skip first 6 header rows (formatting)
                            skipfooter=1,                  # 🚫 Skip footer line if present
                            engine='python',               # 🐍 Use Python engine for flexibility
                            na_values=['NIL', '/0']        # ❓ Replace invalid entries with NaN
                        )
                        dfs.append(df)  # ✅ Add valid DataFrame to the list
                    except ValueError:
                        # ⚠️ Handle files missing required columns gracefully
                        print(f"Skipping {csv_file} in {zip_file.name} - Missing required columns.")

    if dfs:
        return pd.concat(dfs, ignore_index=True)  # 🧩 Combine all valid DataFrames
    else:
        print("No matching CSV files found or required columns missing. ❌")
        return None
```

## 📊📥 4. Load & Clean LTE Hourly Dataset


```python
# 📁 Define path and parameters for LTE CG Hourly data
zip_folder = "D:/Advance_Data_Sets/PPT_Sunset_NW"
keyword_lte = "(LTE_CG_DA)"
required_columns_lte = [
    'Date', 'LTE Cell Group', 'Data Volume (GB)', 
    'L.Traffic.User.VoIP.Avg', 'DL User Thrp (Mbps)',
    'DL User Thrp (Mbps)_NUM', 'DL User Thrp (Mbps)_DEN'
]

# 📦 Load LTE data from ZIP files and rename key columns for standardization
lte_df = load_gsm_from_zip_folder(zip_folder, keyword_lte, required_columns_lte).\
    rename(columns={
        "LTE Cell Group": "LTE_Cell_Group",                            # 🔁 Rename for consistency
        "Data Volume (GB)": "4G Data Traffic (GB)",                    # 📦 4G total traffic
        "L.Traffic.User.VoIP.Avg": "VoLTE Traffic (Erl)",              # 📞 VoLTE traffic
        "DL User Thrp (Mbps)": "4G Overall Throughput (Mbps)"         # 🚀 Downlink throughput
    })

# 🧩 Apply standardized Cell Group (CG) names to LTE data for consistency across regions
lte_df['CG'] = lte_df['LTE_Cell_Group'].replace(standard_cg_mapping)
```

## 📡📶 5. Load & Preprocess LTE L900 Layer Data


```python
# 📦 Load LTE L900 data from ZIP files and rename columns for clarity
keyword_lte_l900 = "(LTE_CG_DA_L900)"
lte_l900_df = load_gsm_from_zip_folder(zip_folder, keyword_lte_l900, required_columns_lte).\
    rename(columns={
        "Data Volume (GB)": "4G Data Volume (GB) L900",                        # 📊 L900 data volume
        "L.Traffic.User.VoIP.Avg": "4G Volte Traffic L900",                   # 📞 VoLTE traffic L900
        "DL User Thrp (Mbps)_NUM": "DL_User_Thrp_Mbps_NUM_L900",              # ➗ Throughput numerator
        "DL User Thrp (Mbps)_DEN": "DL_User_Thrp_Mbps_DEN_L900",              # ➗ Throughput denominator
        "LTE Cell Group": "LTE_Cell_Group_L900",                              # 🔁 Rename for clarity
        "DL User Thrp (Mbps)": "4G Throughput (Mbps) L900"                    # 🚀 Throughput L900
    })


# 🧩 Apply standardized CG names to LTE L900 data to ensure naming consistency
lte_l900_df['CG'] = lte_l900_df['LTE_Cell_Group_L900'].replace(standard_cg_mapping)
```

## 📡⚙️ 6. Load & Clean LTE L1800 Layer Data


```python
# 📦 Load LTE L1800 layer data from ZIPs and rename columns for clarity
keyword_lte_l1800 = "(LTE_CG_DA_L1800)"

lte_l1800_df = load_gsm_from_zip_folder(zip_folder, keyword_lte_l1800, required_columns_lte).\
    rename(columns={
        "Data Volume (GB)": "4G Data Volume (GB) L1800",                         # 📊 L1800 data traffic
        "L.Traffic.User.VoIP.Avg": "4G Volte Traffic L1800",                    # 📞 VoLTE traffic L1800
        "DL User Thrp (Mbps)_NUM": "DL_User_Thrp_Mbps_NUM_L1800",               # ➗ Throughput numerator
        "DL User Thrp (Mbps)_DEN": "DL_User_Thrp_Mbps_DEN_L1800",               # ➗ Throughput denominator
        "LTE Cell Group": "LTE_Cell_Group_L1800",                               # 🔁 Standardized column name
        "DL User Thrp (Mbps)": "4G Throughput (Mbps) L1800"                      # 🚀 Throughput L1800
    })


# 🧩 Apply standardized CG names to LTE L1800 data to maintain uniform naming
lte_l1800_df['CG'] = lte_l1800_df['LTE_Cell_Group_L1800'].replace(standard_cg_mapping)
```

## 📡📶 7. Load & Clean LTE L2100 Layer Data


```python
# 📦 Load LTE L2100 layer data and rename columns for clear identification
keyword_lte_l2100 = "(LTE_CG_DA_L2100)"

lte_l2100_df = load_gsm_from_zip_folder(zip_folder, keyword_lte_l2100, required_columns_lte).\
    rename(columns={
        "Data Volume (GB)": "4G Data Volume (GB) L2100",                         # 📊 L2100 traffic volume
        "L.Traffic.User.VoIP.Avg": "4G Volte Traffic L2100",                    # 📞 VoLTE usage
        "DL User Thrp (Mbps)_NUM": "DL_User_Thrp_Mbps_NUM_L2100",               # ➗ Throughput numerator
        "DL User Thrp (Mbps)_DEN": "DL_User_Thrp_Mbps_DEN_L2100",               # ➗ Throughput denominator
        "LTE Cell Group": "LTE_Cell_Group_L2100",                               # 🏷️ LTE cell group column
        "DL User Thrp (Mbps)": "4G Throughput (Mbps) L2100"                     # 🚀 Downlink throughput
    })


# 🧩 Apply standardized CG names to LTE L2100 data for consistent grouping across datasets
lte_l2100_df['CG'] = lte_l2100_df['LTE_Cell_Group_L2100'].replace(standard_cg_mapping)
```

## 📻📉 8. Load & Clean GSM Voice Data


```python
keyword_gsm = "(GSM_CG_DA)"
required_columns_gsm = ['Date', 'TCH Availability Rate(%)', 'Globel Traffic', 'GCell Group']

gsm_df = load_gsm_from_zip_folder(zip_folder, keyword_gsm, required_columns_gsm).\
    rename(columns={
        "Globel Traffic": "2G Voice Traffic (Erl)",                   # 📞 2G voice traffic volume
        "TCH Availability Rate(%)": "TCH Availability"                # 📶 Channel availability %
    })

# 🧩 Apply standardized CG names to GSM data to ensure uniform grouping across reports
gsm_df['CG'] = gsm_df['GCell Group'].replace(standard_cg_mapping)
```

## 📡📱 9. Load & Clean UMTS (3G) Voice & Data


```python
# 📦 Load UMTS CG Hourly data (3G voice + data) and rename key KPIs
keyword_umts = "(UMTS_CG_DA)"
required_columns_umts = ['Date', 'UCell Group', 'VS.AMR.Erlang.BestCell_SUM', 'Data Volume (GB)']

umts_df = load_gsm_from_zip_folder(zip_folder, keyword_umts, required_columns_umts).\
    rename(columns={
        "VS.AMR.Erlang.BestCell_SUM": "3G Voice Traffic (Erl)",      # 📞 3G Voice traffic (Erlangs)
        "Data Volume (GB)": "3G Data Traffic (GB)"                   # 🌐 3G Data traffic
    })
# 🧩 Apply standardized CG names to UMTS data to maintain consistency across all layers
umts_df['CG'] = umts_df['UCell Group'].replace(standard_cg_mapping)
```

## 🔗🧩 10. Merge All Technology Layers into Unified DataFrame


```python
# 🧱 List of DataFrames to merge in order (UMTS + LTE layers) – GSM is used as the base
dfs_to_merge = [umts_df, lte_df, lte_l900_df, lte_l1800_df, lte_l2100_df]

# 🔗 Merge all DataFrames on ['Date', 'Time', 'CG'] using left joins
merged_df = reduce(
    lambda left, right: pd.merge(left, right, on=['Date', 'CG'], how='left'),
    dfs_to_merge,
    gsm_df.copy()  # ⚙️ Start merging from GSM DataFrame
).fillna(0)        # 🧼 Replace all NaNs with 0 (assume missing = no traffic/activity)
```

## 📊 11. Convert 4G Data Volumes from GB to TB


```python
# 🔄 Convert individual band volumes
merged_df['4G_DataVolume_L900_TB'] = (merged_df['4G Data Volume (GB) L900'] / 1024)  # 🟠 L900 Band
merged_df['4G_DataVolume_L1800_TB'] = (merged_df['4G Data Volume (GB) L1800'] / 1024)  # 🔵 L1800 Band
merged_df['4G_DataVolume_L2100_TB'] = (merged_df['4G Data Volume (GB) L2100'] / 1024)  # 🟣 L2100 Band
# ➕ Calculate total 4G traffic in TB
merged_df['4G_DataVolume_TB'] = (merged_df['4G Data Traffic (GB)'] / 1024)  # 📦 Total 4G Traffic
```

## ➕📊 12. Calculate Total Data Volume & Total Voice Traffic ☎️📈


```python
# ➕📦 Add total data volume by summing 3G and 4G data columns (in GB)
merged_df['Total Data Volume (GB)'] = (merged_df['3G Data Traffic (GB)'] + merged_df['4G Data Traffic (GB)'])
# ➕📞 Add total voice traffic by summing 2G, 3G, and VoLTE voice traffic
merged_df['Total Voice'] = (merged_df['2G Voice Traffic (Erl)'] + merged_df['3G Voice Traffic (Erl)'] + merged_df['VoLTE Traffic (Erl)'])
```

## 🛡️ 13. Handles ZeroDivision & Ensures Total = 100%


```python
# 🧮 User-defined function to calculate band-wise percentage of 4G traffic in TB with % symbol
def adjust_percentages(row):
    total = row['4G_DataVolume_TB']
    
    if total == 0 or pd.isna(total):
        # ⚠️ Avoid division by zero or NaN: return '0%' for all
        return pd.Series({'L1800%': '0%', 'L2100%': '0%', 'L900%': '0%'})
    
    # Step 1️⃣: Raw exact percentages
    parts = {
        'L1800%': (row['4G_DataVolume_L1800_TB'] / total) * 100,
        'L2100%': (row['4G_DataVolume_L2100_TB'] / total) * 100,
        'L900%':  (row['4G_DataVolume_L900_TB']  / total) * 100
    }

    # Step 2️⃣: Floor the percentages
    floored = {k: int(np.floor(v)) for k, v in parts.items()}
    remainder = {k: v - floored[k] for k, v in parts.items()}

    # Step 3️⃣: Distribute the remaining % to top decimal fractions
    diff = 100 - sum(floored.values())
    for k in sorted(remainder, key=remainder.get, reverse=True)[:diff]:
        floored[k] += 1

    # Step 4️⃣: Add % symbol to the final result
    formatted = {k: f"{v}%" for k, v in floored.items()}
    
    return pd.Series(formatted)

#Apply row-wise
merged_df[['L1800%', 'L2100%', 'L900%']] = merged_df.apply(adjust_percentages, axis=1)
```

## 🎯🧮 14. Adjust Rounded KPI Parts to Match Total


```python
def adjust_kpi_parts(df, kpi_mapping):
    """
    🧮 Adjust multiple KPI part columns so that their integer-rounded sum matches the integer-rounded total KPI.

    Parameters:
        df (pd.DataFrame): 📋 The input DataFrame containing KPI totals and parts.
        kpi_mapping (dict): 🗺️ A dictionary where:
            🔑 Keys = total KPI column names (e.g., 'VoLTE Traffic (Erl)')
            📌 Values = list of sub-component KPI column names (e.g., ['Volte L900', 'Volte L1800', ...])

    Returns:
        pd.DataFrame: ✅ A new DataFrame with adjusted parts so that:
                      ➕ Rounded total = sum of rounded parts
    """
    df_adj = df.copy()

    for total_col, part_cols in kpi_mapping.items():
        # 1️⃣ Round total KPI to nearest integer
        total_rounded = df_adj[total_col].round().astype("Int64")

        # 2️⃣ Floor each part (sub-KPI) and calculate remaining decimals (remainders)
        parts_float = df_adj[part_cols]
        parts_floor = parts_float.apply(np.floor)  # ⬇️ Floor values
        remainders = parts_float - parts_floor     # 🔁 Remainder to guide rounding

        # 3️⃣ Calculate how many units need to be added to parts to match the rounded total
        units_needed = (total_rounded - parts_floor.sum(axis=1)).astype(int)

        # 4️⃣ Distribute missing units to the parts with the largest remainders
        adjusted_parts = parts_floor.copy()
        for i in range(len(df_adj)):
            n = units_needed[i]
            if n > 0:
                top_indices = remainders.iloc[i].nlargest(n).index  # 🥇 Get top remainders
                adjusted_parts.loc[i, top_indices] += 1             # ➕ Add 1 to the top parts

        # 5️⃣ Assign adjusted parts and final rounded total to output DataFrame
        df_adj[part_cols] = adjusted_parts.astype("Int64")  # 🔢 Ensure final integer format
        df_adj[total_col] = total_rounded                   # 🧮 Final total

    return df_adj
```

## 🧩📏 15. Apply KPI Part Adjustment for Rounding Consistency


```python
# 🗺️ Define mapping of total KPIs and their corresponding part columns
kpi_mapping = {
    "Total Data Volume (GB)": [                   # 🌐 Total data
        "3G Data Traffic (GB)",
        "4G Data Traffic (GB)"
    ]
}

# 🧮 Apply adjustment so integer-rounded sub-KPIs sum up to the rounded total KPIs
merged_df = adjust_kpi_parts(merged_df, kpi_mapping)
```

## 🧩 16. KPI Part Adjustment: Ensure Parts Sum to Rounded Total (1 Decimal Accuracy)


```python
def adjust_kpi_parts_decimal(df, kpi_mapping):
    """
    🔢 Adjust KPI parts so that their rounded values (to 1 decimal place) sum up to the rounded total KPI.

    Parameters:
        df (pd.DataFrame): 📋 Input DataFrame with total and part KPI values.
        kpi_mapping (dict): 🗺️ Dict with:
            🔑 Keys = total KPI column names
            📌 Values = list of part KPI column names

    Returns:
        pd.DataFrame: ✅ DataFrame with parts rounded to 1 decimal point and their sum matching the rounded total.
    """
    df_adj = df.copy()

    for total_col, part_cols in kpi_mapping.items():
        # 1️⃣ Round total KPI to 1 decimal
        total_rounded = df_adj[total_col].round(1)

        # 2️⃣ Floor parts to 1 decimal and calculate remainders
        parts_float = df_adj[part_cols]
        parts_floor = (parts_float * 10).apply(np.floor) / 10  # ⬇️ Floor to 1 decimal
        remainders = parts_float - parts_floor

        # 3️⃣ Calculate how many 0.1 units are needed
        units_needed = ((total_rounded - parts_floor.sum(axis=1)) * 10).round().astype(int)

        # 4️⃣ Distribute 0.1 to parts with highest remainder
        adjusted_parts = parts_floor.copy()
        for i in range(len(df_adj)):
            n = units_needed[i]
            if n > 0:
                top_indices = remainders.iloc[i].nlargest(n).index
                adjusted_parts.loc[i, top_indices] += 0.1

        # 5️⃣ Round to 1 decimal just in case and assign to output
        df_adj[part_cols] = adjusted_parts.round(1)
        df_adj[total_col] = total_rounded

    return df_adj
```

## 🛠️ 17. Apply KPI Parts Adjustment: Ensure Sub-KPIs Sum to Rounded Total (1 Decimal Precision)


```python
# 🧭 Define which KPIs need part-wise adjustment
kpi_mapping1 = {
    'Total Voice': [  # 📞 Voice traffic components
        "2G Voice Traffic (Erl)",
        "3G Voice Traffic (Erl)",
        "VoLTE Traffic (Erl)"
    ],
    "4G_DataVolume_TB": [  # 💾 4G traffic split by bands
        "4G_DataVolume_L1800_TB",
        "4G_DataVolume_L2100_TB",
        "4G_DataVolume_L900_TB"
    ]
}

# 🔧 Apply decimal adjustment to ensure part KPIs sum to the rounded total
merged_df = adjust_kpi_parts_decimal(merged_df, kpi_mapping1)
```

## 🧮 18. Custom Rounding Function for Flexible Decimal Precision 🎯


```python
# 🧮 Custom Rounding Function for Flexible Decimal Precision
def round_to(x, decimals):
    try:
        return round(float(x), decimals)
    except (ValueError, TypeError):
        return x  # or np.nan
        

# 🛠️ Define rounding precision for each target column
rounding_config = {
    '4G Overall Throughput (Mbps)': 1,
    '4G Throughput (Mbps) L900': 1,
    '4G Throughput (Mbps) L1800': 1,
    '4G Throughput (Mbps) L2100': 1
}

# 🔁 Apply rounding column by column using the config
for col, decimals in rounding_config.items():
    merged_df[col] = merged_df[col].apply(lambda x: round_to(x, decimals))

```

## 🛠️ 19. Function to Zero Out Metrics by Date & CG


```python
def set_columns_to_zero(df, date_values, cg_value, columns_to_zero):
    """
    Set specified columns to zero where Date is in date_values and CG matches cg_value(s).

    Parameters:
    df (pd.DataFrame): The input DataFrame
    date_values (list, set, or pd.Series): Dates to match
    cg_value (str, list, or set): CG value(s) to match
    columns_to_zero (list): Columns to set to zero

    Returns:
    pd.DataFrame: Modified DataFrame
    """
    # Ensure date_values is list-like
    if not isinstance(date_values, (list, set, pd.Series)):
        date_values = [date_values]

    # Ensure cg_value is list-like
    if not isinstance(cg_value, (list, set, pd.Series)):
        cg_value = [cg_value]

    # Create mask
    mask = df['Date'].isin(date_values) & df['CG'].isin(cg_value)
    df.loc[mask, columns_to_zero] = 0
    return df

```

## 🧭 20. Phase-3 | Cleanup of 3G/4G KPIs for Sunset Clusters


```python
updated_df = set_columns_to_zero(
    df=merged_df,
    date_values=[pd.Timestamp('2025-06-28')],
    cg_value=['2G_Sunset_NORTH_B2_SZ','2G_Sunset_N_BANNU_SZ','2G_Sunset_N_KOHAT_SZ','2G_Sunset_N_HANGU_SZ',
             'L9_South_June25_SZ','L9_South_June25_DeraAllAHYAR',
             'GSM_SZ_Center_Region_Phase_03','GSM_Dep_No_BZ_Cluster_15'],
    columns_to_zero=['3G Data Traffic (GB)','3G Voice Traffic (Erl)']
)


updated_df = set_columns_to_zero(
    df=merged_df,
    date_values=[pd.Timestamp('2025-06-27'), pd.Timestamp('2025-06-23')],
    cg_value=['2G_Sunset_NORTH_B2_SZ','2G_Sunset_N_BANNU_SZ','2G_Sunset_N_KOHAT_SZ','2G_Sunset_N_HANGU_SZ',
             'L9_South_June25_SZ','L9_South_June25_DeraAllAHYAR'],
    columns_to_zero=['4G Throughput (Mbps) L900','4G_DataVolume_L900_TB']
)

updated_df = set_columns_to_zero(
    df=merged_df,
    date_values=[pd.Timestamp('2025-06-27')],
    cg_value=['GSM_Dep_No_BZ_Cluster_15'],
    columns_to_zero=['4G Throughput (Mbps) L2100']
)
```

## 🧭 21. Phase-4 | Cleanup of 3G/4G KPIs for Sunset Clusters


```python
updated_df = set_columns_to_zero(
    df=merged_df,
    date_values=[pd.Timestamp('2025-07-10')],
    cg_value=['Sunset_NORTH_B3_SZ','Sunset_N_PESHAWAR_SZ',
    'SZ_Dep_Cluster_9','Sunset_SIALKOT',
    'South_July25_2G_P4_SZ','South_July25_2G_Lark_city'],
    columns_to_zero=['3G Data Traffic (GB)','3G Voice Traffic (Erl)']
)

updated_df = set_columns_to_zero(
    df=merged_df,
    date_values=[pd.Timestamp('2025-07-09')],
    cg_value=['Sunset_NORTH_B3_SZ','Sunset_N_PESHAWAR_SZ',
    'SZ_Dep_Cluster_9','Sunset_SIALKOT',
    'South_July25_2G_P4_SZ','South_July25_2G_Lark_city'],
    columns_to_zero=['4G Throughput (Mbps) L900','4G_DataVolume_L900_TB']
)
```

## 📤 22. Export Processed DataFrame


```python
path = 'D:/Advance_Data_Sets/PPT_Sunset_NW'                    # 📁 Set working directory path for data and templates
os.chdir(path)                                                 # 🔄 Change current working directory to the specified path
merged_df.to_csv('kpi_checking.csv',index=False)
```

## 📊 23. Final Selection of Key KPIs for Reporting & Export 📁


```python
# 🧱 Filter and organize required columns from the merged DataFrame
merged_df1 = merged_df[[
    'Date',                                    # 📅 Date of record
    'CG',                                      # 🗂️ Cell Group
    'Total Voice',                             # ☎️ Total Voice Traffic (all tech)
    "2G Voice Traffic (Erl)",                  # 2️⃣ 2G Voice
    "3G Voice Traffic (Erl)",                  # 3️⃣ 3G Voice
    'VoLTE Traffic (Erl)',                     # 📶 VoLTE Traffic
    'Total Data Volume (GB)',                  # 📦 Total Data in GB
    '3G Data Traffic (GB)',                    # 3️⃣ 3G Data in GB
    '4G Data Traffic (GB)',                    # 4️⃣ 4G Data in GB
    
    # 📐 Converted TB values per 4G band
    '4G_DataVolume_TB',
    '4G_DataVolume_L1800_TB',
    '4G_DataVolume_L2100_TB',
    '4G_DataVolume_L900_TB',
    
    # ⚡ Throughput KPIs
    '4G Overall Throughput (Mbps)',
    '4G Throughput (Mbps) L900',
    '4G Throughput (Mbps) L1800',
    '4G Throughput (Mbps) L2100',

    # 📈 Band-wise Share %
    'L1800%',
    'L2100%',
    'L900%'
]]
```

## 🧠 24. Extract Max Date Rows & Create Dynamic CG-Based Variables


```python
# 📅 Step 1: Filter rows where Date is maximum
result = merged_df1.loc[
    merged_df1['Date'] == merged_df1['Date'].max(), 
    ['Date', 'CG', 'L1800%', 'L2100%', 'L900%']  # 🎯 Select only key columns
]

# 🔁 Step 2: Loop through each CG and create dynamic variable like share_<CG>
for cg in result['CG'].unique():
    # 🧼 Clean CG name to make it a valid Python variable (no spaces or hyphens)
    safe_cg = cg.replace(" ", "_").replace("-", "_")

    # 🏷️ Construct variable name like: share_L9_South_June25_SZ
    var_name = f"share_{safe_cg}"

    # 🧪 Filter DataFrame for this CG and assign it to a dynamic global variable
    globals()[var_name] = result[result['CG'] == cg][['Date', 'CG', 'L1800%', 'L2100%', 'L900%']]
```

## 🧠🔍 25. User-Defined Function: Filter DataFrame by Date, CG List, and Selected Columns 📅🗂️


```python
# 🧠🔍 User-Defined Function: Filter DataFrame by Date, CG List, and Selected Columns 📅🗂️
def filter_by_date_and_columns(df, date_col, start_date, columns_to_keep, rename_dict=None, cg_list=None):
    """
    📅 Filters a DataFrame based on date and CG list, and selects/renames columns.

    Parameters:
    - df: pandas.DataFrame 📊 Input dataset
    - date_col: str 🗓️ Name of the date column
    - start_date: str or datetime ⏳ Minimum date to include
    - columns_to_keep: list[str] ✅ Columns to retain in output
    - rename_dict: dict, optional ✏️ Rename mapping {old_name: new_name}
    - cg_list: list[str], optional 🏷️ Cell Group names to filter

    Returns:
    - pandas.DataFrame 🧾 Filtered and optionally renamed DataFrame
    """
    # 🗓️ Filter rows where the date is on or after the specified start_date
    filtered_df = df[df[date_col] >= pd.to_datetime(start_date)]

    # 🏷️ Apply CG filtering if cg_list is provided and 'CG' column exists
    if cg_list and 'CG' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['CG'].isin(cg_list)]

    # 📑 Retain only columns that exist in the DataFrame
    columns_to_keep = [col for col in columns_to_keep if col in filtered_df.columns]
    filtered_df = filtered_df[columns_to_keep]

    # ✏️ Optionally rename columns using rename_dict
    if rename_dict:
        rename_dict = {k: v for k, v in rename_dict.items() if k in filtered_df.columns}
        filtered_df = filtered_df.rename(columns=rename_dict)

    # ✅ Return the cleaned and filtered DataFrame
    return filtered_df
```

## 🔁📦 26. User-Defined Function: Filter and Combine Multiple CG Groups into One DataFrame 🧾🧩


```python
# 🧠🔍 User-Defined Function: Filter DataFrame by Date, CG List, and Selected Columns 📅🗂️
def filter_by_date_and_columns(df, date_col, start_date, columns_to_keep, rename_dict=None, cg_list=None):
    """
    📅 Filters a DataFrame based on date and CG list, and selects/renames columns.

    Parameters:
    - df: pandas.DataFrame 📊 Input dataset
    - date_col: str 🗓️ Name of the date column
    - start_date: str or datetime ⏳ Minimum date to include
    - columns_to_keep: list[str] ✅ Columns to retain in output
    - rename_dict: dict, optional ✏️ Rename mapping {old_name: new_name}
    - cg_list: list[str], optional 🏷️ Cell Group names to filter

    Returns:
    - pandas.DataFrame 🧾 Filtered and optionally renamed DataFrame
    """
    # 🗓️ Filter rows where the date is on or after the specified start_date
    filtered_df = df[df[date_col] >= pd.to_datetime(start_date)]

    # 🏷️ Apply CG filtering if cg_list is provided and 'CG' column exists
    if cg_list and 'CG' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['CG'].isin(cg_list)]

    # 📑 Retain only columns that exist in the DataFrame
    columns_to_keep = [col for col in columns_to_keep if col in filtered_df.columns]
    filtered_df = filtered_df[columns_to_keep]

    # ✏️ Optionally rename columns using rename_dict
    if rename_dict:
        rename_dict = {k: v for k, v in rename_dict.items() if k in filtered_df.columns}
        filtered_df = filtered_df.rename(columns=rename_dict)

    # ✅ Return the cleaned and filtered DataFrame
    return filtered_df
```

## 📦🔧 27. Configuration & Execution: Filter Multiple CG Groups by Date and KPIs 🗂️📊


```python
# 🔁📦 User-Defined Function: Filter and Combine Multiple CG Groups into One DataFrame 🧾🧩
def filter_multiple_cg_groups(df, date_col, cg_groups, columns_to_keep, rename_dict=None):
    """
    🔁 Applies `filter_by_date_and_columns` to multiple CG groups and combines the results.

    Parameters:
    - df: pandas.DataFrame 📊 Input dataset
    - date_col: str 🗓️ Name of the date column
    - cg_groups: list[dict] 📦 Each group should contain:
        - 'name': str 🏷️ Optional group label (e.g. "Phase 2")
        - 'cg_list': list[str] ➕ Cell Groups to include
        - 'start_date': str or datetime ⏳ Start date for filtering
    - columns_to_keep: list[str] ✅ Columns to retain in each group
    - rename_dict: dict, optional ✏️ Rename mapping {old: new}

    Returns:
    - pandas.DataFrame 🧾 Combined and labeled data from all CG groups
    """
    combined = pd.concat([
        # ✅ Ensure 'CG' is included without duplication
        filter_by_date_and_columns(
            df=df,
            date_col=date_col,
            start_date=group['start_date'],
            columns_to_keep=(columns_to_keep if 'CG' in columns_to_keep else columns_to_keep + ['CG']),
            rename_dict=rename_dict,
            cg_list=group['cg_list']
        ).assign(Group=group.get('name', 'Unknown'))  # 🏷️ Add group name column
        for group in cg_groups
    ], ignore_index=True)  # 🔄 Combine all into a single DataFrame

    return combined
```

## 📦🔧 28. Configuration & Execution: Filter Multiple CG Groups by Date and KPIs 🗂️📊


```python
# 🗃️ Define CG group configurations with labels and filtering start dates
cg_groups = [
    {
        "name": "Phase 3",  
        "cg_list": [
            'L9_South_June25_SZ', 'L9_South_June25_DeraAllAHYAR',
            '2G_Sunset_NORTH_B2_SZ', '2G_Sunset_N_KOHAT_SZ',
            '2G_Sunset_N_HANGU_SZ', '2G_Sunset_N_BANNU_SZ',
            'GSM_BZ_Center_Region_Phase-03',
            'GSM_Dep_Cluster_120', 'GSM_Dep_Cluster_121',
            'GSM_Dep_Cluster_80', 'GSM_Dep_No_BZ_Cluster_15',
            'GSM_Dep_No_BZ_Cluster_19', 'GSM_Dep_No_BZ_Cluster_33',
            'GSM_Dep_No_BZ_Cluster_41', 'GSM_Dep_No_BZ_Cluster_50',
            'GSM_Dep_No_BZ_Cluster_54', 'GSM_Dep_No_BZ_Cluster_64',
            'GSM_SZ_BZ_Center_Region_Phase-03', 'GSM_SZ_Center_Region_Phase_03'
        ],
        "start_date": '2025-06-01'  
    },
    {
        "name": "Phase 2",
        "cg_list": [
            'L9_Jacobabad_Apr25_2G', 'L9_Control_Shikarpur_2G', 'L9_Khuzdar_SZ_2G',
            'L9_Control_MajCitiesBaluchistan_2G', 'L9_Hub_Apr25_2G_SZ',
            'HUB_Control_2G', 'L9_Turbat_Apr25_2G_SZ', 'L2100_Gwadar_2G_City'
        ],
        "start_date": '2025-04-01'
    },
    {
        "name": "Phase 4", 
        "cg_list": [
            'SZ_Dep_Cluster_9', 'Sunset_SAMBRIAL', 'Sunset_SIALKOT',
            'Sunset_N_PESHAWAR_SZ', 'Sunset_NORTH_B3_SZ',
            'South_July25_2G_P4_SZ','South_July25_2G_Lark_city',
        ],
        "start_date": '2025-06-10'
    },

        {
        "name": "Phase 5",
        "cg_list": [
         'SZ_Center_Phase V','DG_KHAN_City',
            'South_July25_22nd Batch_SZ',
            'South_July25_Shahdadkot_city_SZ'
        ],
        "start_date": '2025-06-22'
    },
    {
        "name": "Phase 6",
        "cg_list": [
            'SZ_Center_Phase VI', 
            'DI_KHAN_City'
        ],
        "start_date": '2025-06-29'
    }
]


# 📑 Define the columns to retain for analysis and reporting
columns_to_keep = [
    'Date', 'CG', 'Total Voice', '2G Voice Traffic (Erl)', '3G Voice Traffic (Erl)', 'VoLTE Traffic (Erl)',
    'Total Data Volume (GB)', '3G Data Traffic (GB)', '4G Data Traffic (GB)', '4G_DataVolume_TB',
    '4G_DataVolume_L1800_TB', '4G_DataVolume_L2100_TB', '4G_DataVolume_L900_TB',
    '4G Overall Throughput (Mbps)', '4G Throughput (Mbps) L900',
    '4G Throughput (Mbps) L1800', '4G Throughput (Mbps) L2100'
]

# ✏️ Define column renaming rules for clarity
rename_dict = {
    'Date': 'Date',
    '2G Voice Traffic (Erl)': '2G Voice',
    '3G Voice Traffic (Erl)': '3G Voice',
    'VoLTE Traffic (Erl)': 'VoLTE',
    'Total Voice': 'Total Voice',
    '3G Data Traffic (GB)': '3G Data Volume (GB)',
    '4G Data Traffic (GB)': '4G Data Volume (GB)',
    'Total Data Volume (GB)': 'Data Volume'
}

# 🚀 Execute the filtering across all CG groups and combine the results
filtered_all = filter_multiple_cg_groups(
    df=merged_df1,                     # 📊 Input DataFrame
    date_col='Date',                   # 📅 Column to filter by date
    cg_groups=cg_groups,               # 📦 CG group configurations
    columns_to_keep=columns_to_keep,   # ✅ Selected columns
    rename_dict=rename_dict            # ✏️ Rename mapping
)
```

## 📊🧱 29. CG-wise Split: Voice, Data, Layer Volumes, and Throughput


```python
# 🔍📊 User-Defined Function: Extract Views for a Single CG from DataFrame
def extract_cg_views(df, cg_name):
    """
    📊 Returns structured views (dataframes) for a given CG name:
    - Voice traffic
    - Total data
    - Layer-wise 4G data volume
    - Layer-wise 4G throughput

    Parameters:
    - df: pandas.DataFrame with filtered/renamed columns (must include 'CG' column)
    - cg_name: str ➡️ Name of the CG to extract

    Returns:
    - dict of DataFrames: {'voice': ..., 'data': ..., 'volume': ..., 'throughput': ...}
    """
    # 🎯 Filter data for the specified CG
    df_cg = df[df['CG'] == cg_name]

    # 🗣️ Voice Traffic View
    df_voice = df_cg[['Date', '2G Voice', '3G Voice', 'VoLTE', 'Total Voice']]

    # 📶 Data Traffic View
    df_data = df_cg[['Date', '3G Data Volume (GB)', '4G Data Volume (GB)', 'Data Volume']]

    # 📦 Layer-wise 4G Data Volume View
    df_volume = (
        df_cg[['Date', '4G_DataVolume_L900_TB', '4G_DataVolume_L1800_TB',
               '4G_DataVolume_L2100_TB', '4G_DataVolume_TB']]
        .rename(columns={
            '4G_DataVolume_L900_TB': 'L900',
            '4G_DataVolume_L1800_TB': 'L1800',
            '4G_DataVolume_L2100_TB': 'L2100',
            '4G_DataVolume_TB': 'Total'
        })
        [['Date', 'L1800', 'L2100', 'L900', 'Total']]
    )

    # 🚀 Layer-wise 4G Throughput View
    df_throughput = (
        df_cg[['Date', '4G Overall Throughput (Mbps)', '4G Throughput (Mbps) L900',
               '4G Throughput (Mbps) L1800', '4G Throughput (Mbps) L2100']]
        .rename(columns={
            '4G Overall Throughput (Mbps)': 'Overall',
            '4G Throughput (Mbps) L900': 'L900',
            '4G Throughput (Mbps) L1800': 'L1800',
            '4G Throughput (Mbps) L2100': 'L2100'
        })
        [['Date', 'Overall', 'L1800', 'L2100', 'L900']]
    )

    # 📦 Return all structured views in a dictionary
    return {
        'voice': df_voice,
        'data': df_data,
        'volume': df_volume,
        'throughput': df_throughput
    }

```

## 🧬 30. Apply Dynamic View Creation for Each CG and Rename 'Date' ➡️ 'Row Labels'


```python
# 📦 Step 1: Extract unique CGs from filtered DataFrame
cg_list = filtered_all['CG'].unique().tolist()

# 🗂️ Initialize view dictionaries
voice_views = {}
data_views = {}
volume_views = {}
throughput_views = {}

# 🔁 Extract views for each CG using the updated function
for cg in cg_list:
    result = extract_cg_views(filtered_all, cg)
    voice_views[cg] = result['voice']
    data_views[cg] = result['data']
    volume_views[cg] = result['volume']
    throughput_views[cg] = result['throughput']

# 🧠 Step 2: Dynamically assign views to variables and rename 'Date' ➡️ 'Row Labels'
for cg in cg_list:
    var_suffix = cg.lower().replace('-', '_').replace(' ', '_')

    # 🎯 Assign to dynamic variable names
    locals()[f"voice_{var_suffix}"] = voice_views[cg]
    locals()[f"data_{var_suffix}"] = data_views[cg]
    locals()[f"volume_{var_suffix}"] = volume_views[cg]
    locals()[f"throughput_{var_suffix}"] = throughput_views[cg]

    # ✏️ Rename 'Date' column to 'Row Labels'
    for view_type in ['voice', 'data', 'volume', 'throughput']:
        view_var = locals()[f"{view_type}_{var_suffix}"]
        if 'Date' in view_var.columns:
            view_var.rename(columns={'Date': 'Row Labels'}, inplace=True)
```

## 🔍📦 31. Identify All CG-Specific DataFrames by Type (Voice, Data, Volume, Throughput)


```python
[v for v in locals() if v.startswith('voice_')]         # 🗣️ Find all local variables related to voice traffic
```




    ['voice_views',
     'voice_2g_sunset_north_b2_sz',
     'voice_2g_sunset_n_bannu_sz',
     'voice_2g_sunset_n_hangu_sz',
     'voice_2g_sunset_n_kohat_sz',
     'voice_gsm_bz_center_region_phase_03',
     'voice_gsm_dep_cluster_120',
     'voice_gsm_dep_cluster_121',
     'voice_gsm_dep_cluster_80',
     'voice_gsm_dep_no_bz_cluster_15',
     'voice_gsm_dep_no_bz_cluster_19',
     'voice_gsm_dep_no_bz_cluster_33',
     'voice_gsm_dep_no_bz_cluster_41',
     'voice_gsm_dep_no_bz_cluster_50',
     'voice_gsm_dep_no_bz_cluster_54',
     'voice_gsm_dep_no_bz_cluster_64',
     'voice_gsm_sz_bz_center_region_phase_03',
     'voice_gsm_sz_center_region_phase_03',
     'voice_l9_south_june25_sz',
     'voice_l9_south_june25_deraallahyar',
     'voice_l9_jacobabad_apr25_2g',
     'voice_l9_hub_apr25_2g_sz',
     'voice_l9_turbat_apr25_2g_sz',
     'voice_l9_khuzdar_sz_2g',
     'voice_l9_control_majcitiesbaluchistan_2g',
     'voice_l9_control_shikarpur_2g',
     'voice_l2100_gwadar_2g_city',
     'voice_hub_control_2g',
     'voice_sunset_n_peshawar_sz',
     'voice_sunset_north_b3_sz',
     'voice_sz_dep_cluster_9',
     'voice_sunset_sambrial',
     'voice_sunset_sialkot',
     'voice_south_july25_2g_p4_sz',
     'voice_south_july25_2g_lark_city',
     'voice_sz_center_phase_v',
     'voice_dg_khan_city',
     'voice_south_july25_22nd_batch_sz',
     'voice_south_july25_shahdadkot_city_sz',
     'voice_sz_center_phase_vi',
     'voice_di_khan_city']




```python
[v for v in locals() if v.startswith('data_')]          # 📶 Find all local variables related to total data traffic
```




    ['data_views',
     'data_2g_sunset_north_b2_sz',
     'data_2g_sunset_n_bannu_sz',
     'data_2g_sunset_n_hangu_sz',
     'data_2g_sunset_n_kohat_sz',
     'data_gsm_bz_center_region_phase_03',
     'data_gsm_dep_cluster_120',
     'data_gsm_dep_cluster_121',
     'data_gsm_dep_cluster_80',
     'data_gsm_dep_no_bz_cluster_15',
     'data_gsm_dep_no_bz_cluster_19',
     'data_gsm_dep_no_bz_cluster_33',
     'data_gsm_dep_no_bz_cluster_41',
     'data_gsm_dep_no_bz_cluster_50',
     'data_gsm_dep_no_bz_cluster_54',
     'data_gsm_dep_no_bz_cluster_64',
     'data_gsm_sz_bz_center_region_phase_03',
     'data_gsm_sz_center_region_phase_03',
     'data_l9_south_june25_sz',
     'data_l9_south_june25_deraallahyar',
     'data_l9_jacobabad_apr25_2g',
     'data_l9_hub_apr25_2g_sz',
     'data_l9_turbat_apr25_2g_sz',
     'data_l9_khuzdar_sz_2g',
     'data_l9_control_majcitiesbaluchistan_2g',
     'data_l9_control_shikarpur_2g',
     'data_l2100_gwadar_2g_city',
     'data_hub_control_2g',
     'data_sunset_n_peshawar_sz',
     'data_sunset_north_b3_sz',
     'data_sz_dep_cluster_9',
     'data_sunset_sambrial',
     'data_sunset_sialkot',
     'data_south_july25_2g_p4_sz',
     'data_south_july25_2g_lark_city',
     'data_sz_center_phase_v',
     'data_dg_khan_city',
     'data_south_july25_22nd_batch_sz',
     'data_south_july25_shahdadkot_city_sz',
     'data_sz_center_phase_vi',
     'data_di_khan_city']




```python
[v for v in locals() if v.startswith('volume_')]        # 🧱 Find all local variables related to 4G layer-wise data volume
```




    ['volume_views',
     'volume_2g_sunset_north_b2_sz',
     'volume_2g_sunset_n_bannu_sz',
     'volume_2g_sunset_n_hangu_sz',
     'volume_2g_sunset_n_kohat_sz',
     'volume_gsm_bz_center_region_phase_03',
     'volume_gsm_dep_cluster_120',
     'volume_gsm_dep_cluster_121',
     'volume_gsm_dep_cluster_80',
     'volume_gsm_dep_no_bz_cluster_15',
     'volume_gsm_dep_no_bz_cluster_19',
     'volume_gsm_dep_no_bz_cluster_33',
     'volume_gsm_dep_no_bz_cluster_41',
     'volume_gsm_dep_no_bz_cluster_50',
     'volume_gsm_dep_no_bz_cluster_54',
     'volume_gsm_dep_no_bz_cluster_64',
     'volume_gsm_sz_bz_center_region_phase_03',
     'volume_gsm_sz_center_region_phase_03',
     'volume_l9_south_june25_sz',
     'volume_l9_south_june25_deraallahyar',
     'volume_l9_jacobabad_apr25_2g',
     'volume_l9_hub_apr25_2g_sz',
     'volume_l9_turbat_apr25_2g_sz',
     'volume_l9_khuzdar_sz_2g',
     'volume_l9_control_majcitiesbaluchistan_2g',
     'volume_l9_control_shikarpur_2g',
     'volume_l2100_gwadar_2g_city',
     'volume_hub_control_2g',
     'volume_sunset_n_peshawar_sz',
     'volume_sunset_north_b3_sz',
     'volume_sz_dep_cluster_9',
     'volume_sunset_sambrial',
     'volume_sunset_sialkot',
     'volume_south_july25_2g_p4_sz',
     'volume_south_july25_2g_lark_city',
     'volume_sz_center_phase_v',
     'volume_dg_khan_city',
     'volume_south_july25_22nd_batch_sz',
     'volume_south_july25_shahdadkot_city_sz',
     'volume_sz_center_phase_vi',
     'volume_di_khan_city']




```python
[v for v in locals() if v.startswith('throughput_')]    # 🚀 Find all local variables related to 4G throughput KPIs
```




    ['throughput_views',
     'throughput_2g_sunset_north_b2_sz',
     'throughput_2g_sunset_n_bannu_sz',
     'throughput_2g_sunset_n_hangu_sz',
     'throughput_2g_sunset_n_kohat_sz',
     'throughput_gsm_bz_center_region_phase_03',
     'throughput_gsm_dep_cluster_120',
     'throughput_gsm_dep_cluster_121',
     'throughput_gsm_dep_cluster_80',
     'throughput_gsm_dep_no_bz_cluster_15',
     'throughput_gsm_dep_no_bz_cluster_19',
     'throughput_gsm_dep_no_bz_cluster_33',
     'throughput_gsm_dep_no_bz_cluster_41',
     'throughput_gsm_dep_no_bz_cluster_50',
     'throughput_gsm_dep_no_bz_cluster_54',
     'throughput_gsm_dep_no_bz_cluster_64',
     'throughput_gsm_sz_bz_center_region_phase_03',
     'throughput_gsm_sz_center_region_phase_03',
     'throughput_l9_south_june25_sz',
     'throughput_l9_south_june25_deraallahyar',
     'throughput_l9_jacobabad_apr25_2g',
     'throughput_l9_hub_apr25_2g_sz',
     'throughput_l9_turbat_apr25_2g_sz',
     'throughput_l9_khuzdar_sz_2g',
     'throughput_l9_control_majcitiesbaluchistan_2g',
     'throughput_l9_control_shikarpur_2g',
     'throughput_l2100_gwadar_2g_city',
     'throughput_hub_control_2g',
     'throughput_sunset_n_peshawar_sz',
     'throughput_sunset_north_b3_sz',
     'throughput_sz_dep_cluster_9',
     'throughput_sunset_sambrial',
     'throughput_sunset_sialkot',
     'throughput_south_july25_2g_p4_sz',
     'throughput_south_july25_2g_lark_city',
     'throughput_sz_center_phase_v',
     'throughput_dg_khan_city',
     'throughput_south_july25_22nd_batch_sz',
     'throughput_south_july25_shahdadkot_city_sz',
     'throughput_sz_center_phase_vi',
     'throughput_di_khan_city']



## 🗂️📊 32. Load PowerPoint template for chart updates 


```python
path = 'D:/Advance_Data_Sets/PPT_Sunset_NW'                    # 📁 Set working directory path for data and templates
os.chdir(path)                                                 # 🔄 Change current working directory to the specified path
prs_phase3 = Presentation('template_phase3.pptx')              # 📊 Load PowerPoint template for Phase-3
prs_phase4 = Presentation('template_phase4.pptx')              # 📊 Load PowerPoint template for Phase-4
```

## 🧩 33. User-Defined Function to Update PowerPoint Chart with DataFrame


```python
def update_chart(prs, slide_index, shape_index, dataframe):
    # 📊 Initialize a new chart data object
    chart_data = CategoryChartData()
    # 🏷️ Set categories (typically X-axis labels) using the first column
    chart_data.categories = dataframe['Row Labels']
    # 📈 Add each series (data columns) to the chart, skipping the category column
    for col in dataframe.columns[1:]:  # Skip 'Row Labels'
        chart_data.add_series(col, dataframe[col])
    # 📄 Access the specific slide by index
    slide = prs.slides[slide_index]
    # 🔳 Access the specific shape on the slide (assumed to be a chart)
    chart_shape = slide.shapes[shape_index]
    # 🧠 Extract the chart object from the shape
    chart = chart_shape.chart
    # 🔄 Replace existing chart data with the new data
    chart.replace_data(chart_data)
```

## 📝 34. Update Table Cells in PowerPoint Slide 📊


```python
def update_table_cells(prs, slide_index, table_index, updates, force_formatting=False):
    """
    Update table cell text in a specific table on a specific slide.

    Parameters:
    - prs : Presentation object
    - slide_index : int, slide number (0-based)
    - table_index : int, table number on slide (0-based)
    - updates : list of tuples (row, col, new_text)
    - force_formatting : bool, if True apply Calibri 11pt bold white center formatting
    """

    slide = prs.slides[slide_index]

    # Get all tables on the slide
    tables = [shape.table for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.TABLE]

    if table_index >= len(tables):
        print(f"No table at index {table_index} on slide {slide_index}")
        return

    table = tables[table_index]

    for row, col, new_text in updates:
        try:
            cell = table.cell(row, col)
            paragraph = cell.text_frame.paragraphs[0]

            if paragraph.runs:
                run = paragraph.runs[0]
                run.text = new_text
            else:
                # If no runs exist, set text normally
                cell.text = new_text
                paragraph = cell.text_frame.paragraphs[0]
                run = paragraph.runs[0]

            if force_formatting:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)  # white
                paragraph.alignment = PP_ALIGN.CENTER

        except IndexError:
            print(f"Cell at row {row}, col {col} does not exist.")
```

## 🔁 35. Reusing User-Defined Function to Update Slide Charts

### 35.1 Phase-3 PPT Charts Update


```python
# 📈 North - Phase-3 Charts

# 📊 Data: 2G Sunset - North B2 SZ & Bannu SZ (Slide 2)
update_chart(prs=prs_phase3, slide_index=2, shape_index=1, dataframe=data_2g_sunset_north_b2_sz)    # 🗺️ North B2 SZ - 2G Sunset Data
update_chart(prs=prs_phase3, slide_index=2, shape_index=0, dataframe=data_2g_sunset_n_bannu_sz)     # 🗺️ Bannu SZ - 2G Sunset Data

# 📊 Data: 2G Sunset - Kohat SZ & Hangu SZ (Slide 3)
update_chart(prs=prs_phase3, slide_index=3, shape_index=1, dataframe=data_2g_sunset_n_kohat_sz)     # 🗺️ Kohat SZ - 2G Sunset Data
update_chart(prs=prs_phase3, slide_index=3, shape_index=0, dataframe=data_2g_sunset_n_hangu_sz)     # 🗺️ Hangu SZ - 2G Sunset Data

# 📦 Volume: 2G Sunset - North B2 SZ & Bannu SZ (Slide 4)
update_chart(prs=prs_phase3, slide_index=4, shape_index=5, dataframe=volume_2g_sunset_north_b2_sz)  # 🧮 North B2 SZ - Volume Data
update_chart(prs=prs_phase3, slide_index=4, shape_index=4, dataframe=volume_2g_sunset_n_bannu_sz)   # 🧮 Bannu SZ - Volume Data

# 📦 Volume: 2G Sunset - Kohat SZ & Hangu SZ (Slide 5)
update_chart(prs=prs_phase3, slide_index=5, shape_index=5, dataframe=volume_2g_sunset_n_kohat_sz)   # 🧮 Kohat SZ - Volume Data
update_chart(prs=prs_phase3, slide_index=5, shape_index=4, dataframe=volume_2g_sunset_n_hangu_sz)   # 🧮 Hangu SZ - Volume Data

# 🚀 Throughput: 2G Sunset - North B2 SZ & Bannu SZ (Slide 6) just change this one
update_chart(prs=prs_phase3, slide_index=6, shape_index=3, dataframe=throughput_2g_sunset_north_b2_sz)  # 🚀 North B2 SZ - Throughput
update_chart(prs=prs_phase3, slide_index=6, shape_index=0, dataframe=throughput_2g_sunset_n_bannu_sz)   # 🚀 Bannu SZ - Throughput

# 🚀 Throughput: 2G Sunset - Kohat SZ & Hangu SZ (Slide 7)
update_chart(prs=prs_phase3, slide_index=7, shape_index=1, dataframe=throughput_2g_sunset_n_kohat_sz)   # 🚀 Kohat SZ - Throughput
update_chart(prs=prs_phase3, slide_index=7, shape_index=0, dataframe=throughput_2g_sunset_n_hangu_sz)   # 🚀 Hangu SZ - Throughput

# 📞 Voice: 2G Sunset - North B2 SZ & Bannu SZ (Slide 8)
update_chart(prs=prs_phase3, slide_index=8, shape_index=4, dataframe=voice_2g_sunset_north_b2_sz)       # 📞 North B2 SZ - Voice
update_chart(prs=prs_phase3, slide_index=8, shape_index=3, dataframe=voice_2g_sunset_n_bannu_sz)        # 📞 Bannu SZ - Voice

# 📞 Voice: 2G Sunset - Kohat SZ & Hangu SZ (Slide 9)
update_chart(prs=prs_phase3, slide_index=9, shape_index=1, dataframe=voice_2g_sunset_n_kohat_sz)        # 📞 Kohat SZ - Voice
update_chart(prs=prs_phase3, slide_index=9, shape_index=0, dataframe=voice_2g_sunset_n_hangu_sz)        # 📞 Hangu SZ - Voice
```


```python
# 📈 South - Phase-3 Charts

# 📊 Data: L9 South June25 SZ & Dera Allahyar (Slide 11)
update_chart(prs=prs_phase3, slide_index=11, shape_index=1, dataframe=data_l9_south_june25_sz)             # 🗺️ L9 South SZ - 2G Sunset Data
update_chart(prs=prs_phase3, slide_index=11, shape_index=0, dataframe=data_l9_south_june25_deraallahyar)   # 🗺️ Dera Allahyar - 2G Sunset Data

# 📦 Volume: L9 South June25 SZ & Dera Allahyar (Slide 12)
update_chart(prs=prs_phase3, slide_index=12, shape_index=5, dataframe=volume_l9_south_june25_sz)           # 🧮 L9 South SZ - Volume
update_chart(prs=prs_phase3, slide_index=12, shape_index=4, dataframe=volume_l9_south_june25_deraallahyar) # 🧮 Dera Allahyar - Volume

# 🚀 Throughput: L9 South June25 SZ & Dera Allahyar (Slide 13)
update_chart(prs=prs_phase3, slide_index=13, shape_index=4, dataframe=throughput_l9_south_june25_sz)       # 🚀 L9 South SZ - Throughput
update_chart(prs=prs_phase3, slide_index=13, shape_index=3, dataframe=throughput_l9_south_june25_deraallahyar) # 🚀 Dera Allahyar - Throughput

# 📞 Voice: L9 South June25 SZ & Dera Allahyar (Slide 14)
update_chart(prs=prs_phase3, slide_index=14, shape_index=4, dataframe=voice_l9_south_june25_sz)            # 📞 L9 South SZ - Voice
update_chart(prs=prs_phase3, slide_index=14, shape_index=3, dataframe=voice_l9_south_june25_deraallahyar)  # 📞 Dera Allahyar - Voice
```


```python
# 🏙️ Center - Phase-3 Charts

# 📊 Data: GSM SZ Center Region & Dep No BZ Cluster 15 (Slide 16)
update_chart(prs=prs_phase3, slide_index=16, shape_index=1, dataframe=data_gsm_sz_center_region_phase_03)        # 🗺️ Center Region - GSM Data
update_chart(prs=prs_phase3, slide_index=16, shape_index=0, dataframe=data_gsm_dep_no_bz_cluster_15)             # 🗺️ Dep No BZ Cluster 15 - GSM Data

# 📦 Volume: GSM SZ Center Region & Dep No BZ Cluster 15 (Slide 17)
update_chart(prs=prs_phase3, slide_index=17, shape_index=4, dataframe=volume_gsm_sz_center_region_phase_03)      # 🧮 Center Region - Volume
update_chart(prs=prs_phase3, slide_index=17, shape_index=3, dataframe=volume_gsm_dep_no_bz_cluster_15)           # 🧮 Dep No BZ Cluster 15 - Volume

# 🚀 Throughput: GSM SZ Center Region & Dep No BZ Cluster 15 (Slide 18)
update_chart(prs=prs_phase3, slide_index=18, shape_index=3, dataframe=throughput_gsm_sz_center_region_phase_03)  # 🚀 Center Region - Throughput
update_chart(prs=prs_phase3, slide_index=18, shape_index=2, dataframe=throughput_gsm_dep_no_bz_cluster_15)       # 🚀 Dep No BZ Cluster 15 - Throughput

# 📞 Voice: GSM SZ Center Region & Dep No BZ Cluster 15 (Slide 19)
update_chart(prs=prs_phase3, slide_index=19, shape_index=3, dataframe=voice_gsm_sz_center_region_phase_03)       # 📞 Center Region - Voice
update_chart(prs=prs_phase3, slide_index=19, shape_index=2, dataframe=voice_gsm_dep_no_bz_cluster_15)            # 📞 Dep No BZ Cluster 15 - Voice
```

### 35.2 Phase-4 PPT Charts Update


```python
# 📈 North - Phase-4 Charts

# 📊 Data:
update_chart(prs=prs_phase4, slide_index=2, shape_index=1, dataframe=data_sunset_north_b3_sz)    
update_chart(prs=prs_phase4, slide_index=2, shape_index=0, dataframe=data_sunset_n_peshawar_sz)     
# 📦 Volume:
update_chart(prs=prs_phase4, slide_index=3, shape_index=4, dataframe=volume_sunset_north_b3_sz)  
update_chart(prs=prs_phase4, slide_index=3, shape_index=3, dataframe=volume_sunset_n_peshawar_sz)   

# 🚀 Throughput: 
update_chart(prs=prs_phase4, slide_index=4, shape_index=3, dataframe=throughput_sunset_north_b3_sz)  
update_chart(prs=prs_phase4, slide_index=4, shape_index=2, dataframe=throughput_sunset_n_peshawar_sz)   

# 📞 Voice: 
update_chart(prs=prs_phase4, slide_index=5, shape_index=3, dataframe=voice_sunset_north_b3_sz)       
update_chart(prs=prs_phase4, slide_index=5, shape_index=2, dataframe=voice_sunset_n_peshawar_sz)        
```


```python
# 📈 Center - Phase-4 Charts

# 📊 Data: 
update_chart(prs=prs_phase4, slide_index=7, shape_index=1, dataframe=data_sz_dep_cluster_9)    
update_chart(prs=prs_phase4, slide_index=7, shape_index=0, dataframe=data_sunset_sialkot)     

# 📦 Volume:
update_chart(prs=prs_phase4, slide_index=8, shape_index=4, dataframe=volume_sz_dep_cluster_9)  
update_chart(prs=prs_phase4, slide_index=8, shape_index=3, dataframe=volume_sunset_sialkot)   

# 🚀 Throughput: 
update_chart(prs=prs_phase4, slide_index=9, shape_index=3, dataframe=throughput_sz_dep_cluster_9)  
update_chart(prs=prs_phase4, slide_index=9, shape_index=2, dataframe=throughput_sunset_sialkot)  

# 📞 Voice: 
update_chart(prs=prs_phase4, slide_index=10, shape_index=3, dataframe=voice_sz_dep_cluster_9)      
update_chart(prs=prs_phase4, slide_index=10, shape_index=2, dataframe=voice_sunset_sialkot)        
```


```python
# 📈 South - Phase-4 Charts

# 📊 Data: 
update_chart(prs=prs_phase4, slide_index=12, shape_index=1, dataframe=data_south_july25_2g_p4_sz)    
update_chart(prs=prs_phase4, slide_index=12, shape_index=0, dataframe=data_south_july25_2g_lark_city)     

# 📦 Volume:
update_chart(prs=prs_phase4, slide_index=13, shape_index=4, dataframe=volume_south_july25_2g_p4_sz)  
update_chart(prs=prs_phase4, slide_index=13, shape_index=3, dataframe=volume_south_july25_2g_lark_city)   

# 🚀 Throughput: 
update_chart(prs=prs_phase4, slide_index=14, shape_index=3, dataframe=throughput_south_july25_2g_p4_sz)  
update_chart(prs=prs_phase4, slide_index=14, shape_index=2, dataframe=throughput_south_july25_2g_lark_city)  

# 📞 Voice: 
update_chart(prs=prs_phase4, slide_index=15, shape_index=0, dataframe=voice_south_july25_2g_p4_sz)      
update_chart(prs=prs_phase4, slide_index=15, shape_index=3, dataframe=voice_south_july25_2g_lark_city)       
```

## 🔁 36. Reusing User-Defined Function to Update Tables

### 36.1 Phase-3 PPT Table Update


```python
# 📋 North - Phase-3 Table Updates

# 🧮 Share Table: Sunset NORTH B2 SZ
update_table_cells(prs=prs_phase3, slide_index=4, table_index=0, 
    updates=[
        (0, 1, f"{share_2G_Sunset_NORTH_B2_SZ.iloc[0, 2]}"),  # 📌 L1800%
        (1, 1, f"{share_2G_Sunset_NORTH_B2_SZ.iloc[0, 3]}"),  # 📌 L2100%
        (2, 1, f"{share_2G_Sunset_NORTH_B2_SZ.iloc[0, 4]}")   # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: BANNU SZ 
update_table_cells(prs=prs_phase3, slide_index=4, table_index=1, 
    updates=[
        (0, 1, f"{share_2G_Sunset_N_BANNU_SZ.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_2G_Sunset_N_BANNU_SZ.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_2G_Sunset_N_BANNU_SZ.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: KOHAT SZ 
update_table_cells(prs=prs_phase3, slide_index=5, table_index=0, 
    updates=[
        (0, 1, f"{share_2G_Sunset_N_KOHAT_SZ.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_2G_Sunset_N_KOHAT_SZ.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_2G_Sunset_N_KOHAT_SZ.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: HANGU SZ 
update_table_cells(prs=prs_phase3, slide_index=5, table_index=1, 
    updates=[
        (0, 1, f"{share_2G_Sunset_N_HANGU_SZ.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_2G_Sunset_N_HANGU_SZ.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_2G_Sunset_N_HANGU_SZ.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)
```


```python
# 📋 South - Phase-3 Table Updates

# 🧮 Share Table: South June25 SZ 
update_table_cells(
    prs=prs_phase3, 
    slide_index=12, 
    table_index=0, 
    updates=[
        (0, 1, f"{share_L9_South_June25_SZ.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_L9_South_June25_SZ.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_L9_South_June25_SZ.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: Dera Allahyar 
update_table_cells(
    prs=prs_phase3, 
    slide_index=12, 
    table_index=1, 
    updates=[
        (0, 1, f"{share_L9_South_June25_DeraAllAHYAR.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_L9_South_June25_DeraAllAHYAR.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_L9_South_June25_DeraAllAHYAR.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)
```


```python
# 📋 Center - Phase-3 Table Updates

# 🧮  GSM SZ Center Region Phase-03 
update_table_cells(
    prs=prs_phase3, 
    slide_index=17, 
    table_index=0, 
    updates=[
        (0, 1, f"{share_GSM_SZ_Center_Region_Phase_03.iloc[0, 2]}"),  # 📌 L1800%
        (1, 1, f"{share_GSM_SZ_Center_Region_Phase_03.iloc[0, 3]}"),  # 📌 L2100%
        (2, 1, f"{share_GSM_SZ_Center_Region_Phase_03.iloc[0, 4]}")   # 📌 L900%
    ],
    force_formatting=True
)

# 🧮 Share Table: GSM Dep No BZ Cluster 15 
update_table_cells(
    prs=prs_phase3, 
    slide_index=17, 
    table_index=1, 
    updates=[
        (0, 1, f"{share_GSM_Dep_No_BZ_Cluster_15.iloc[0, 2]}"),       # 📌 L1800%
        (1, 1, f"{share_GSM_Dep_No_BZ_Cluster_15.iloc[0, 3]}"),       # 📌 L2100%
        (2, 1, f"{share_GSM_Dep_No_BZ_Cluster_15.iloc[0, 4]}")        # 📌 L900%
    ],
    force_formatting=True
)
```

### 36.2 Phase-4 PPT Table Update


```python
# 📋 North - Phase-4 Table Updates

# 🧮 Share Table: North Service Zone
update_table_cells(prs=prs_phase4, slide_index=3, table_index=0, 
    updates=[
        (0, 1, f"{share_Sunset_NORTH_B3_SZ.iloc[0, 2]}"),  # 📌 L1800%
        (1, 1, f"{share_Sunset_NORTH_B3_SZ.iloc[0, 3]}"),  # 📌 L2100%
        (2, 1, f"{share_Sunset_NORTH_B3_SZ.iloc[0, 4]}")   # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: Peshawar
update_table_cells(prs=prs_phase4, slide_index=3, table_index=1, 
    updates=[
        (0, 1, f"{share_Sunset_N_PESHAWAR_SZ.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_Sunset_N_PESHAWAR_SZ.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_Sunset_N_PESHAWAR_SZ.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)
```


```python
# 📋 Center - Phase-4 Table Updates

# 🧮 Share Table: Center Service Zone
update_table_cells(prs=prs_phase4, slide_index=8, table_index=0, 
    updates=[
        (0, 1, f"{share_SZ_Dep_Cluster_9.iloc[0, 2]}"),  # 📌 L1800%
        (1, 1, f"{share_SZ_Dep_Cluster_9.iloc[0, 3]}"),  # 📌 L2100%
        (2, 1, f"{share_SZ_Dep_Cluster_9.iloc[0, 4]}")   # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: Sialkot
update_table_cells(prs=prs_phase4, slide_index=8, table_index=1, 
    updates=[
        (0, 1, f"{share_Sunset_SIALKOT.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_Sunset_SIALKOT.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_Sunset_SIALKOT.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)

```


```python
# 📋 South - Phase-4 Table Updates

# 🧮 Share Table: Center Service Zone
update_table_cells(prs=prs_phase4, slide_index=13, table_index=0, 
    updates=[
        (0, 1, f"{share_South_July25_2G_P4_SZ.iloc[0, 2]}"),  # 📌 L1800%
        (1, 1, f"{share_South_July25_2G_P4_SZ.iloc[0, 3]}"),  # 📌 L2100%
        (2, 1, f"{share_South_July25_2G_P4_SZ.iloc[0, 4]}")   # 📌 L900%
    ], 
    force_formatting=True
)

# 🧮 Share Table: Peshawar
update_table_cells(prs=prs_phase4, slide_index=13, table_index=1, 
    updates=[
        (0, 1, f"{share_South_July25_2G_Lark_city.iloc[0, 2]}"),   # 📌 L1800%
        (1, 1, f"{share_South_July25_2G_Lark_city.iloc[0, 3]}"),   # 📌 L2100%
        (2, 1, f"{share_South_July25_2G_Lark_city.iloc[0, 4]}")    # 📌 L900%
    ], 
    force_formatting=True
)
```

## 💾 36. Save Presentation


```python
prs_phase3.save('UMTS_Sunset_Phase_3.pptx')                       # 💾 Save the updated PowerPoint presentation to file
prs_phase4.save('UMTS_Sunset_Phase_4.pptx')                       # 💾 Save the updated PowerPoint presentation to file
```


```python
%reset -f
```

## 🧹 37. PowerPoint Cleanup - Automated


```python
import os
import glob
import win32com.client

# Folder path and file filter
ppt_folder = r"D:\Advance_Data_Sets\PPT_Sunset_NW"
ppt_files = glob.glob(os.path.join(ppt_folder, "UMTS_Sunset*.pptx"))

# Launch PowerPoint
ppt = win32com.client.Dispatch("PowerPoint.Application")
# Do not set ppt.Visible = False

for ppt_path in ppt_files:
    print(f"Processing: {ppt_path}")
    try:
        presentation = ppt.Presentations.Open(ppt_path, WithWindow=False)

        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasChart:
                    try:
                        chart = shape.Chart
                        workbook = chart.ChartData.Workbook
                        sheet = workbook.Worksheets(1)
                        used_range = sheet.UsedRange

                        for row in range(1, used_range.Rows.Count + 1):
                            for col in range(1, used_range.Columns.Count + 1):
                                value = sheet.Cells(row, col).Value
                                if value == 0:
                                    sheet.Cells(row, col).Value = None  # or "=NA()"

                    except Exception as chart_err:
                        print(f"  Chart error on Slide {slide.SlideIndex}: {chart_err}")

        presentation.Save()
        presentation.Close()

    except Exception as file_err:
        print(f"  Error opening {ppt_path}: {file_err}")

# Gracefully quit PowerPoint
try:
    if ppt:
        ppt.Quit()
except Exception as quit_err:
    print(f"⚠️ Could not quit PowerPoint: {quit_err}")

print("✅ All matching PowerPoint files processed.")
```

    Processing: D:\Advance_Data_Sets\PPT_Sunset_NW\UMTS_Sunset_Phase_3.pptx
    Processing: D:\Advance_Data_Sets\PPT_Sunset_NW\UMTS_Sunset_Phase_4.pptx
    ✅ All matching PowerPoint files processed.
    


```python
%reset -f
```

import os                      # 📁 For interacting with the operating system (file paths, directory checks, etc.)
import zipfile                 # 🗜️ For handling ZIP archive files
import numpy as np             # 🔢 For numerical operations and arrays
import pandas as pd            # 🐼 For data manipulation and analysis
from pathlib import Path       # 🛤️ For object-oriented file path handling
from datetime import time      # ⏰ For working with time objects (e.g., filtering by hour)
from functools import reduce   # 🔄 For function-based reduction (e.g., combining multiple DataFrames)
from pandas import Timestamp   # 📅 For handling timestamps specifically in pandas

from pptx.util import Inches                   # 📐 For specifying size/dimensions of shapes or images in inches
from pptx import Presentation                  # 📊 For creating and editing PowerPoint presentations
from pptx.util import Pt                       # 🔠 For setting font size in points
from pptx.enum.shapes import MSO_SHAPE_TYPE    # 🧩 For identifying and handling different shape types in slides
from pptx.dml.color import RGBColor            # 🎨 For setting custom RGB colors for text/shapes
from pptx.enum.text import PP_ALIGN            # 📏 For setting paragraph alignment (left, center, right, etc.)
from pptx.chart.data import CategoryChartData  # 📈 For providing data to category charts (like bar/column charts)

path = 'D:/Advance_Data_Sets/PPT_Sunset_NW'                    # 📁 Set working directory path for data and templates
os.chdir(path)                                                 # 🔄 Change current working directory to the specified path

prs_phase4 = Presentation('template_phase3.pptx')              # 📊 Load PowerPoint template for Phase-4

# Choose the slide number (e.g., second slide)
slide = prs_phase4 .slides[19]  # 0-based index

# Loop through and print info about each shape
for idx, shape in enumerate(slide.shapes):
    shape_type = shape.shape_type
    name = getattr(shape, "name", "No name")
    print(f"Index: {idx}, Type: {shape_type}, Name: {name}, Has Chart: {shape.has_chart}")
