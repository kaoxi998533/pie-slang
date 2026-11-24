d = {"a": 1, "b": 2}
d["c"] = 3
print(d, len(d), d["b"], "a" in d, "z" in d)
print(list(d.keys()), list(d.values()), list(d.items()))
print(d.get("a"), d.get("z"), d.get("z", 0))
d.update({"d": 4}, e=5)
print(d)
print(d.pop("a"), d.pop("zz", None), d)
d.setdefault("f", []).append(1)
d.setdefault("f", []).append(2)
print(d["f"])
del d["b"]
print(d)
for k, v in d.items():
    print(k, "->", v)
sq = {x: x * x for x in range(5)}
print(sq)
print(dict(zip("abc", [1, 2, 3])), dict([("x", 1)]), dict(a=1, b=2))
counts = {}
for w in "the cat the dog the end".split():
    counts[w] = counts.get(w, 0) + 1
print(counts)
print(sorted(counts, key=counts.get, reverse=True)[0])
print({1: "int", 1.0: "float", True: "bool"})
k = {(1, 2): "tuple key"}
print(k[(1, 2)], k[1, 2])
s = {3, 1, 2, 3}
print(sorted(s), len(s), 2 in s)
s.add(10); s.discard(1); s.discard(99)
print(sorted(s))
a, b = {1, 2, 3}, {2, 3, 4}
print(sorted(a | b), sorted(a & b), sorted(a - b), sorted(a ^ b))
print(sorted(a.union(b)), sorted(a.intersection(b)), a.issubset({1, 2, 3, 4}), a.isdisjoint({9}))
print(set(), sorted(set("hello")), {x % 3 for x in range(10)} == {0, 1, 2})
e = {}
print(e, len(e), bool(e))
d2 = {"x": [1, 2], "y": {"z": 1}}
print(d2["y"]["z"], d2 == {"y": {"z": 1}, "x": [1, 2]})
print(d.popitem(), d)
m = {"b": 1, "a": 2}
print(sorted(m.items()), max(m, key=lambda k: m[k]))
try:
    print(m["missing"])
except KeyError as ex:
    print("KeyError", ex)
