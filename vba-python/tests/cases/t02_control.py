for i in range(5):
    if i == 1:
        continue
    if i == 4:
        break
    print("i =", i)
else:
    print("not printed")

for i in range(3):
    pass
else:
    print("for-else ran", i)

n = 0
while n < 10:
    n += 3
else:
    print("while-else", n)

while True:
    n -= 1
    if n < 5:
        break
print("n", n)

def classify(v):
    if v < 0:
        return "neg"
    elif v == 0:
        return "zero"
    elif v < 10:
        return "small"
    else:
        return "big"

print([classify(v) for v in (-3, 0, 5, 50)])

total = 0
for i in range(1, 4):
    for j in range(1, 4):
        if j > i:
            break
        total += i * j
print("total", total)

for i in range(10, 0, -3):
    print(i, end=" ")
print()
for ch in "abc":
    print(ch * 2, end=",")
print()
i = 0
while i < 3: i += 1
print(i)
if i == 3: print("one-line if")
for k in []:
    print("never")
print(list(range(0)), list(range(2, 8, 2)), list(range(5, 0, -2)))
