# calc tools

def mapping_calc(mark, mapping, ceil_score, ceil_max):
    floor_score = int(mark)
    ceil_score = min(floor_score + 1, ceil_score)
    floor_value = mapping.get(floor_score, 0)
    ceil_value = mapping.get(ceil_score, ceil_max)
    return floor_value + (ceil_value - floor_value) * (mark - floor_score)