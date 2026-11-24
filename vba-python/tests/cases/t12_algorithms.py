def sieve(n):
    is_p = [True] * (n + 1)
    is_p[0] = is_p[1] = False
    for i in range(2, int(n ** 0.5) + 1):
        if is_p[i]:
            for j in range(i * i, n + 1, i):
                is_p[j] = False
    return [i for i, p in enumerate(is_p) if p]
print(sieve(60))

def quicksort(xs):
    if len(xs) <= 1:
        return xs
    pivot, *rest = xs
    return quicksort([x for x in rest if x < pivot]) + [pivot] + quicksort([x for x in rest if x >= pivot])
print(quicksort([5, 2, 9, 1, 5, 6, 0, -3]))

def bsearch(xs, target):
    lo, hi = 0, len(xs) - 1
    while lo <= hi:
        mid = (lo + hi) // 2
        if xs[mid] == target:
            return mid
        if xs[mid] < target:
            lo = mid + 1
        else:
            hi = mid - 1
    return -1
print(bsearch(list(range(0, 100, 3)), 42), bsearch([1, 2, 3], 7))

def queens(n):
    sols = []
    def place(row, cols, d1, d2, cur):
        if row == n:
            sols.append(cur[:])
            return
        for c in range(n):
            if c not in cols and row - c not in d1 and row + c not in d2:
                cur.append(c)
                place(row + 1, cols | {c}, d1 | {row - c}, d2 | {row + c}, cur)
                cur.pop()
    place(0, set(), set(), set(), [])
    return sols
q = queens(6)
print(len(q), q[0])

def matmul(A, B):
    return [[sum(a * b for a, b in zip(row, col)) for col in zip(*B)] for row in A]
print(matmul([[1, 2], [3, 4]], [[5, 6], [7, 8]]))

text = "the quick brown fox jumps over the lazy dog the end"
freq = {}
for w in text.split():
    freq[w] = freq.get(w, 0) + 1
top = sorted(freq.items(), key=lambda kv: (-kv[1], kv[0]))[:3]
print(top)

def hanoi(n, a="A", b="B", c="C", moves=None):
    if moves is None:
        moves = []
    if n:
        hanoi(n - 1, a, c, b, moves)
        moves.append(a + c)
        hanoi(n - 1, b, a, c, moves)
    return moves
print(len(hanoi(10)), hanoi(2))

def gcd(a, b):
    while b:
        a, b = b, a % b
    return a
print(gcd(1071, 462), [n for n in range(2, 30) if all(n % d for d in range(2, n))])

memo = {}
def collatz(n):
    if n == 1:
        return 1
    if n in memo:
        return memo[n]
    r = 1 + collatz(n // 2 if n % 2 == 0 else 3 * n + 1)
    memo[n] = r
    return r
best = max(range(1, 400), key=collatz)
print(best, collatz(best))

grid = ["#..#", "....", ".##."]
print(sum(row.count("#") for row in grid), ["".join(r) for r in zip(*grid)])
def flatten(x):
    return [z for y in x for z in (flatten(y) if isinstance(y, list) else [y])]
print(flatten([1, [2, [3, 4], 5], [[6]]]))
print(" ".join(f"{i}x{j}={i*j}" for i in range(1, 4) for j in range(1, 4) if i <= j))
acc = 0
for i in range(3000):
    acc += i % 7
print(acc)
