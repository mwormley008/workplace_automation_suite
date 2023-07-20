import itertools

numbers = [1180, 1120, 3739, 3330, 2698]

combinations = []
for r in range(1, len(numbers) + 1):
    combinations.extend(itertools.combinations(numbers, r))

for comb in combinations:
    comb_sum = sum(comb)
    if comb_sum == 8660:
        print(f"Combination: {comb}  Sum: {comb_sum}")
else:
    print("No other matching combination(s) found.")
