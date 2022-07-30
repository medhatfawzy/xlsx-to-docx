numbers = ( x for x in range(1, 500))


def coordination(codes:tuple) -> str:
    string = []
    for num in codes:
        string.append(str(num).zfill(2))
        
    string.append(str(next(numbers)).zfill(3))
    return "-".join(string)
