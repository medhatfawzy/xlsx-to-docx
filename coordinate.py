numbers = ( x for x in range(1, 200))


def coordination(codes:tuple) -> str:
    string = []
    for i, num in enumerate(codes):
        string.append(str(num).zfill(i+2))
    string.append(str(next(numbers)).zfill(4))
    return "-".join(string)
