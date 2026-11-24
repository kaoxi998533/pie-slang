import math
from math import sqrt, pi as PI
import collections
from collections import Counter, defaultdict
from itertools import permutations, combinations, product, chain, accumulate
from functools import reduce, lru_cache
import string

print(math.sqrt(16), sqrt(2) > 1.414, round(PI, 5), math.floor(-2.5), math.ceil(2.1), math.trunc(-3.7))
print(math.factorial(10), math.gcd(12, 18), math.lcm(4, 6), math.isqrt(99), math.comb(5, 2), math.perm(5, 2))
print(round(math.sin(math.pi / 2), 6), round(math.cos(0), 6), round(math.log(math.e), 6), math.log10(1000), math.log2(8))
print(math.hypot(3, 4), math.fabs(-2), math.pow(2, 3), round(math.degrees(math.pi), 6), math.isclose(0.1 + 0.2, 0.3))
print(math.prod([1, 2, 3, 4]), round(math.atan2(1, 1) * 4, 6), math.fsum([0.1] * 10), math.inf > 1e308)
c = Counter("mississippi")
print(c.most_common(2), c["s"], c["z"], sorted(c.items()))
c2 = Counter(["a", "b", "a"])
c2.update(["b", "b"])
print(c2, c2.total() if hasattr(c2, "total") else 4)
dd = defaultdict(list)
for word in ["apple", "avocado", "banana", "blueberry", "cherry"]:
    dd[word[0]].append(word)
print(dict(dd), len(dd))
di = defaultdict(int)
for ch in "hello":
    di[ch] += 1
print(sorted(di.items()))
print(list(permutations([1, 2, 3])), list(permutations("ab", 2)))
print(list(combinations(range(4), 2)), list(product("ab", repeat=2)), list(product([1, 2], "xy")))
print(list(chain([1, 2], (3,), "ab")), list(accumulate([1, 2, 3, 4])), list(accumulate([1, 2, 3], lambda a, b: a * b)))
print(reduce(lambda a, b: a * b, range(1, 6)), reduce(lambda a, b: a + b, [], 0))
@lru_cache(maxsize=None)
def fib(n):
    return n if n < 2 else fib(n - 1) + fib(n - 2)
print(fib(40), fib(30))
print(string.ascii_lowercase[:5], string.digits, "a" in string.ascii_letters)
od = collections.OrderedDict()
od["x"] = 1
od["y"] = 2
print(list(od.keys()))
import random
random.seed(1)
r = random.random()
print(0 <= r < 1, 1 <= random.randint(1, 6) <= 6, random.choice([7]) == 7, sorted(random.sample(range(5), 5)))
lst = list(range(10))
random.shuffle(lst)
print(sorted(lst) == list(range(10)))
