def greet(name, greeting="Hello", *rest, sep=", ", **opts):
    s = greeting + sep + name
    if rest:
        s += " " + str(rest)
    if opts:
        s += " " + str(sorted(opts.items()))
    return s

print(greet("Ann"))
print(greet("Bob", "Hi"))
print(greet("Cy", "Yo", 1, 2, sep="; ", x=1, a=2))
print(greet(greeting="Hey", name="Di"))

def fact(n):
    return 1 if n <= 1 else n * fact(n - 1)
print(fact(10), fact(12))

def make_counter():
    count = 0
    def inc(step=1):
        nonlocal count
        count += step
        return count
    return inc

c1 = make_counter()
c1(); c1()
print(c1(10))

total = 0
def add_total(v):
    global total
    total += v
add_total(5); add_total(7)
print("total", total)

square = lambda x: x * x
print(square(7), (lambda a, b=2: a ** b)(3), list(map(lambda t: t[0] + t[1], [(1, 2), (3, 4)])))

def apply(f, *args, **kw):
    return f(*args, **kw)
print(apply(max, 3, 8, 1), apply(greet, "Ed", greeting="Bye"))

def deco(fn):
    def wrapper(*a, **k):
        return "<" + str(fn(*a, **k)) + ">"
    return wrapper

@deco
def plus(a, b):
    return a + b
print(plus(2, 3))

def adders():
    return [lambda x, i=i: x + i for i in range(3)]
print([f(10) for f in adders()])

def kwonly(a, *, b=5):
    return a * b
print(kwonly(2), kwonly(2, b=3))

def no_return():
    pass
print(no_return())

def fib_gen(n):
    out = []
    a, b = 0, 1
    while len(out) < n:
        out.append(a)
        a, b = b, a + b
    return out
print(fib_gen(15))
print(plus.__name__ if hasattr(plus, "__name__") else "?", fact.__name__)
def outer():
    x = "outer"
    def inner():
        return x
    x = "changed"
    return inner()
print(outer())
