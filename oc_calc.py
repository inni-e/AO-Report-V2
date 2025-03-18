# Module for OC Calculator
from util import mapping_calc

def oc_eng_mapping(raw):
    mapping = {0:0,
               1:18.24, 
               2:21.71,
               3:24.45,
               4:26.82,
               5:28.99,
               6:31.1,
               7:33.2,
               8:34,
               9:35.38,
               10:37.76,
               11:40.5,
               12:43.98,
               13:49.32,
               14:50}
    if raw in mapping:
        return mapping[raw]
    else:
        mapping_calc(raw, mapping, 14, 50)

def oc_maths_mapping(raw):
    mapping = {0:0,
               1:18.00,
               2:20.58,
               3:22.55,
               4:24.64,
               5:26.43,
               6:28.12,
               7:29.77,
               8:31.42,
               9:33.1,
               10:34.85,
               11:36.75,
               12:38.87,
               13:41.36,
               14:44.59,
               15:47.61,
               16:49.66,
               17:49.77,
               18:50}
    if raw in mapping:
        return mapping[raw]
    else:
        mapping_calc(raw, mapping, 18, 50)

def oc_thinking_mapping(raw):
    mapping = {0:0, 
               1:19.51,
               2:24.54,
               3:26.59,
               4:28.5,
               5:30.32,
               6:32.08,
               7:32.1,
               8:33.84,
               9:35.64, 
               10:35.96,
               11:37.5,
               12:39.53,
               13:40.53,
               14:41.78,
               15:42.78,
               16:44.43,
               17:48,
               18:50}
    if raw in mapping:
        return mapping[raw]
    else:
        mapping_calc(raw, mapping, 18, 50)

def oc_calc(marks_list):
    eng = (marks_list[0]/25)*14
    maths = (marks_list[1]/35)*18
    thinking = (marks_list[2]/30)*18
    
    # scaling factors
    mean_scaled = 3.410917837399822
    e_scale = 0.79523116
    m_scale = 0.80474513
    t_scale = 0.70451434
    
    # calculate scaled marks (using linear interpolation)
    # english
    eng_mark = oc_eng_mapping(eng)*e_scale
    
    # maths
    maths_mark = oc_maths_mapping(maths)*m_scale
    
    # thinking skils
    thinking_mark = oc_thinking_mapping(thinking)*t_scale
    
    mark = mean_scaled + eng_mark + maths_mark + thinking_mark
    
    if mark > 100:
        pad = (120 - mark) / 4
        mark += pad

    return mark