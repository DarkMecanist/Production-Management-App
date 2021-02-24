# TEST LOOPING FROM THE MIDDLE

# from itertools import cycle
# import math
#
# job_time = 60
#
# time_list = cycle(['6:00', '6:15', '6:30', '6:45', '7:00', '7:15', '7:30', '7:45', '8:00'])
#
# num_cells = math.ceil(job_time / 15) + 1
#
# print(num_cells)
#
# for _ in range(num_cells):
#     c = next(time_list)
#
# print(c)

from itertools import cycle

def return_next_possible_time(time_string, time_list):
    valid_time = False
    iterator_time_list = cycle(time_list)

    while valid_time == False:
        try:
            time_list.index(next(iterator_time_list))
        except ValueError:
            continue

