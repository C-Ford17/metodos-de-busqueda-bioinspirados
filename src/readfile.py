
def readfile(file):
    coords = []

    if 'tiny.csv' in file:
        with open(file, "r") as f:
            for line in f:
                x, y = line.split(',')
                coords.append((float(x), float(y)))
    elif 'coord.txt' in file:
        with open(file, "r") as f:
            for line in f:
                x, y = line.split(' ')
                coords.append((float(x), float(y)))
    elif 'TSP51.txt' in file:
        with open(file, "r") as f:
            for line in f:
                _, x, y = line.split(' ')
                coords.append((float(x), float(y)))
    elif 'uscap50.txt' in file:
        with open(file, "r") as f:
            for line in f:
                x, y = line.split(' ')
                coords.append((float(x), float(y)))
    
    return coords