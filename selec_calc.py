# Module for Selective Calculator
from util import mapping_calc
# Mapping functions
def sttc_eng_mapping(mark):
    mapping = {0:0,
               1:3.29,
               2:5.15,
               3:7.01,
               4:8.87,
               5:10.73,
               6:12.59,
               7:14.45,
               8:16.42,
               9:18.17,
               10:20.03,
               11:21.89,
               12:23.75,
               13:25.21,
               14:26.38,
               15:27.53,
               16:28.66,
               17:29.78,
               18:30.91,
               19:32.04,
               20:33.19,
               21:34.36,
               22:35.58,
               23:36.86,
               24:38.21,
               25:39.66,
               26:41.24,
               27:43.02,
               28:45.09,
               29:47.63,
               30:50}
    if mark in mapping:
        return mapping[mark]
    else:
        mapping_calc(mark, mapping, 30, 50)

def sttc_maths_mapping(mark):
    mapping = {0:0,
               1:2.04,
               2:4.2,
               3:6.36,
               4:8.52,
               5:10.68,
               6:12.84,
               7:15,
               8:17.16,
               9:19.32,
               10:21.48,
               11:22.98,
               12:23.88,
               13:24.78,
               14:25.68,
               15:26.58,
               16:27.48,
               17:28.38,
               18:29.28,
               19:29.98,
               20:30.7,
               21:31.43,
               22:32.93,
               23:33.72,
               24:34.55,
               25:35.66,
               26:36.36,
               27:36.5,
               28:37.38,
               29:38.5,
               30:39.77,
               31:40.77,
               32:41.25,
               33:43.07,
               34:45.52,
               35:49.5}
    if mark in mapping:
        return mapping[mark]
    else:
        mapping_calc(mark, mapping, 35, 49.5)

def sttc_thinking_mapping(mark):
    mapping = {0:0,
               1:3.94,
               2:5.12,
               3:6.3,
               4:7.48,
               5:8.66,
               6:9.84,
               7:11.02,
               8:12.2,
               9:13.38,
               10:14.56,
               11:15.73,
               12:16.92,
               13:18.1,
               14:19.28,
               15:20.46,
               16:21.64,
               17:22.82,
               18:24,
               19:25.18,
               20:26.36,
               21:27.54,
               22:28.72,
               23:29.9,
               24:31.06,
               25:32.21,
               26:33.36,
               27:34.5,
               28:35.66,
               29:36.82,
               30:38,
               31:39.2,
               32:40.43,
               33:41.7,
               34:43.01,
               35:44.38,
               36:45.83,
               37:46,
               38:47.5,
               39:49.03,
               40:50}
    if mark in mapping:
        return mapping[mark]
    else:
        mapping_calc(mark, mapping, 40, 50)

def sttc_writing_mapping(mark):
    mapping = {0:0,
               1:4,
               2:8,
               3:11.2,
               4:13.6,
               5:16,
               6:18.4,
               7:20,
               8:21.6,
               9:23.2,
               10:24.8,
               11:25.6,
               12:27.2,
               13:28.8,
               14:30.4,
               15:32,
               16:33.6,
               17:35.2,
               18:36.8,
               19:38.4,
               20:40,
               21:42,
               22:44,
               23:46,
               24:48,
               25:50}
    if mark in mapping:
        return mapping[mark]
    else:
        mapping_calc(mark, mapping, 25, 50)

# selective calc function
def selective_calc(marks_list):
    # scale factors
    e_scale = 0.54578
    m_scale = 0.56031
    t_scale = 0.79127
    w_scale = 0.41064
    
    # change raw mark
    eng_mark = sttc_eng_mapping(marks_list[0]) * e_scale
    maths_mark = sttc_maths_mapping(marks_list[1]) * m_scale
    thinking_mark = sttc_thinking_mapping(marks_list[2]) * t_scale
    writing_mark = sttc_writing_mapping(int(marks_list[3])) * w_scale
    
    final_mark = calc_helper(eng_mark, maths_mark, thinking_mark, writing_mark)
    return(final_mark)


def calc_helper(eng_mark, maths_mark, thinking_mark, writing_mark):

    # New 2024 weightings are 25% all
    reweighted_thinking_skills = 0.25/0.35
    reweighted_writing_skills = 0.25/0.15

    mark = eng_mark + maths_mark + thinking_mark * reweighted_thinking_skills + writing_mark * reweighted_writing_skills
    mark = min(120, mark)
    mark = max(0, mark)

    theoretical_max = 115.12
    if mark >= 110:
        mark = (mark * 120) / theoretical_max

    return mark