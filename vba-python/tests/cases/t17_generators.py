def count_up(n):
    i = 0
    while i < n:
        yield i
        i += 1

def evens(xs):
    for x in xs:
        if x % 2 == 0:
            yield x * 10

def chained():
    yield "start"
    yield from count_up(3)
    yield from ["x", "y"]
    return

print(list(count_up(5)), sum(count_up(101)), list(evens(range(7))), list(chained()))
for v in count_up(3):
    print("v", v)
g = count_up(2)
print(next(g), next(g), next(g, "end"))

def fib():
    a, b = 0, 1
    for _ in range(10):
        yield a
        a, b = b, a + b
print(list(fib()), max(fib()), [x for x in fib() if x > 5])

class Tree:
    def __init__(self, v, kids=()):
        self.v, self.kids = v, list(kids)
    def __iter__(self):
        yield self.v
        for k in self.kids:
            yield from k
t = Tree(1, [Tree(2, [Tree(3)]), Tree(4)])
print(list(t), 3 in t)
def pairs(d):
    for k in sorted(d):
        yield k, d[k]
print(dict(pairs({"b": 2, "a": 1})))
