import itertools

numbers = [750, 1200, 750, 6000, 3040, 600, 900, -5745]

combinations = []
for r in range(1, len(numbers) + 1):
    combinations.extend(itertools.combinations(numbers, r))

for comb in combinations:
    comb_sum = sum(comb)
    print(comb_sum)
    if comb_sum == 3064:
        print(f"Combination: {comb}  Sum: {comb_sum}")
else:
    print("No other matching combination(s) found.")
