import math

def compute_frametimes(frametime_data, metric, value=-1.0):
    cdef double result = 0
    cdef double current_total = 0

    if metric == "Max":
        result = frametime_data["min"]
    elif metric == "Avg":
        result = frametime_data["sum"] / frametime_data["len"]
    elif metric == "Min":
        result = frametime_data["max"]
    elif metric == "Percentile" and value > -1:
        result = frametime_data["frametimes"][math.ceil(value / 100 * frametime_data["len"]) - 1]
    elif metric == "Lows" and value > -1:
        for present in frametime_data["frametimes"]:
            current_total += present
            if current_total >= value / 100 * frametime_data["sum"]:
                result = present
                break
    elif metric == "STDEV":
        mean = frametime_data["sum"] / frametime_data["len"]
        dev = [x - mean for x in frametime_data["frametimes"]]
        dev2 = [x * x for x in dev]
        result = math.sqrt(sum(dev2) / frametime_data["len"])
    return result