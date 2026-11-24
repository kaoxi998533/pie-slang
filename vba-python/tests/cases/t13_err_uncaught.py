def f(x):
    return g(x) + 1
def g(x):
    return {"a": 1}[x]
print("before")
f("zz")
print("not reached")
