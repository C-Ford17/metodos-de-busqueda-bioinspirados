from multiprocessing import Pool
from timeit import default_timer
from copy import deepcopy
from random import choice
from statistics import variance, mean
from met_simulated_annealing import SimAnneal


def thread_proc(sa: SimAnneal):
    start = default_timer()
    result = sa.anneal()
    end = default_timer()
    time = end - start
    return time, *result


def parallel_sa(coords: list, params: dict, nexe: int, nthreads: int):
    sa = SimAnneal(
        coords, 
        params['[temperatura_inicial]'],
        params['[tasa_enfriamiento]'],
        params['[temperatura_final]'],
        params['[sa_max_iter]'])
    
    pool = Pool(nthreads)
    params = [deepcopy(sa) for _ in range(nexe)]
    durations, paths, distances, iterations, fitness_list = zip(*pool.map(thread_proc, params))
    
    min_distance = min(distances)
    max_distance = max(distances)
    avg_distance = mean(distances)
    var_distance = variance(distances)
    avg_time = mean(durations)
    avg_iter = mean(iterations)
    
    min_path = paths[distances.index(min_distance)]

    return (
        min_distance,
        max_distance,
        avg_distance,
        var_distance,
        avg_time,
        avg_iter,
        min_path,
        choice(fitness_list)
    )
    