def forever(n):
    return forever(n + 1)
try:
    forever(0)
except RecursionError as e:
    print("recursion caught")
print(1 / 0)
