a = [5, 3, 8, 1]
a.append(7)
a.insert(0, 9)
print(a, len(a), a[0], a[-1], a[1:3], a[::-1], a[::2])
a.sort()
print(a)
a.sort(reverse=True)
print(a)
print(a.pop(), a.pop(0), a)
a.remove(5)
print(a, a.index(8), a.count(3), 3 in a, 4 not in a)
a.extend([1, 1])
a += [2]
print(a)
a[1:3] = ["x", "y", "z"]
print(a)
del a[0]
del a[-2:]
print(a)
b = a.copy(); b.reverse(); print(b, a)
m = [[0] * 3 for _ in range(3)]
m[1][1] = 5
print(m)
print([x * x for x in range(10) if x % 2 == 0])
print([(x, y) for x in range(3) for y in range(3) if x != y])
print(sorted(["banana", "Apple", "cherry"], key=str.lower))
print(sorted(["bb", "a", "ccc"], key=len, reverse=True))
words = ["pear", "fig", "apple", "kiwi"]
print(sorted(words, key=lambda w: (len(w), w)))
print([1, 2] * 3, [0] + [1], list("abc"), list(range(3)))
nested = [1, [2, [3, [4]]]]
print(nested[1][1][1][0])
s = [3, 1, 2]
print(sorted(s), s, max(s), min(s), sum(s))
print(list(reversed([1, 2, 3])), list(enumerate("ab")), list(zip([1, 2, 3], "ab")))
x = [1, 2, 3, 4, 5, 6]
print(x[1:-1], x[-3:], x[:100], x[5:1:-1], x[::-2])
x[::2] = [0, 0, 0]
print(x)
print([] == [], [1] == [1.0], [1, [2]] == [1, [2]])
e = []
print(e, bool(e), len(e))
c = [1, 2, 3]
d = c
d.append(4)
print(c)
t = (1, 2, 3)
print(t, t[1], t[-1:], len(t), (1,), (), t + (4,), t * 2, t.index(2), t.count(1))
print(tuple([1, 2]), list((3, 4)))
