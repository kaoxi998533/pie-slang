a, b = 1, 2
a, b = b, a
print(a, b)
(x, y), z = (1, 2), 3
print(x, y, z)
first, *rest = [1, 2, 3, 4]
print(first, rest)
*init, last = "abcd"
print(init, last)
p, *mid, q = range(5)
print(p, mid, q)
for i, (k, v) in enumerate({"a": 1, "b": 2}.items()):
    print(i, k, v)
pairs = [(1, "one"), (2, "two")]
nums, names = zip(*pairs)
print(nums, names)
print(*[1, 2, 3], sep="-")
def f(a, b, c): return a + b * c
args = [1, 2]
print(f(*args, 3), f(**{"a": 1, "b": 2, "c": 3}))
print([*range(3), *"ab"], {**{"a": 1}, "b": 2})
matrix = [[1, 2, 3], [4, 5, 6]]
print([list(col) for col in zip(*matrix)])
data = [("x", 1, 2.0), ("y", 3, 4.0)]
for name, i, fl in data:
    print(name, i * 2, fl / 2)
try:
    a, b = [1, 2, 3]
except ValueError as e:
    print(e)
try:
    a, b, c = [1, 2]
except ValueError as e:
    print(e)
t = 1,
print(t, len(t))
s = 1, 2, 3
print(s)
d = {}
d["k"], d["j"] = 5, 6
print(d)
lst = [0, 0]
lst[0], lst[1] = "a", "b"
print(lst)
if (n := len(lst)) > 1:
    print("walrus", n)
