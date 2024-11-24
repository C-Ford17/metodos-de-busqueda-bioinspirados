from multiprocessing import Pool
from statistics import mean, variance
from timeit import default_timer
from random import choice
from copy import deepcopy
from math import sqrt


from met_ant_colony import ACO, Graph
import matplotlib.pyplot as plt


def distance(city1: dict, city2: dict):
    return sqrt((city1[0] - city2[0]) ** 2 + (city1[1] - city2[1]) ** 2)


def exec_aco(aco: ACO):
    start = default_timer()
    path, distance, fitness_list, n_gen = aco.solve()
    end = default_timer()
    duration = end - start
    return duration, path, distance, n_gen, fitness_list

def parallel_aco(coords: list, params: dict, nexe: int, nthreads: int):
    cost_matrix = []
    rank = len(coords)
    for i in range(rank):
        row = []
        for j in range(rank):
            row.append(distance(coords[i], coords[j]))
        cost_matrix.append(row)
    
    graph = Graph(cost_matrix, rank)
    config = (
        params['[n_ants]'],
        params['[aco_max_iter]'],
        params['[ro]'],
        params['[alpha]'],
        params['[betha]'],
        1,
        3
    )

    pool = Pool(nthreads)
    args = [ACO(*config, graph=deepcopy(graph)) for _ in range(nexe)]
    durations, paths, distances, generations, fitness_list = zip(*pool.map(exec_aco, args)) 

    min_distance = min(distances)
    max_distance = max(distances)
    avg_distance = mean(distances)
    var_distance = variance(distances)
    avg_duration = mean(durations)
    avg_iter = mean(generations)
    min_path = paths[distances.index(min_distance)]
    return (
        min_distance,
        max_distance,
        avg_distance,
        var_distance,
        avg_duration,
        avg_iter,
        min_path,
        choice(fitness_list)
    )
