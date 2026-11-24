价格 = {"苹果": 3, "香蕉": 2}
总计 = sum(价格.values())
print(f"总计: {总计}", sorted(价格))
d = {"key": [1, 2, 3]}
print(f"{d['key'][1]} {len(d['key'])} {'yes' if d else 'no'} {d!r}")
print(f"{'ab':*^8}|{3.14159:10.4f}|{-42:>6}|{12:x}")
width, prec = 10, 3
print(f"{2.718281828:{width}.{prec}f}|")
s = "Python"
print(s[1:4], s[-3:], s[::-1], s.lower().count("o"))
x = [1, 2, 3]
y = x[:]
y.append(4)
print(x, y, x is y)
if __name__ == "__main__":
    print("main guard ok")
matrix = [[i * j for j in range(4)] for i in range(4)]
print("\n".join(" ".join(f"{v:2}" for v in row) for row in matrix))
def outer(n):
    def helper(k):
        return k * n
    return [helper(i) for i in range(4)]
print(outer(3))
a = [3, 1, 2]
a.sort(key=lambda v: -v)
print(a, sorted(a, key=abs), list(map(str.upper, ["a", "b"])))
print(all(isinstance(v, int) for v in a), any(v > 2 for v in a))
print(", ".join(map(str, range(5))), "%d%%" % 50, "a\tb".expandtabs(4))
print(int("101", 2), int("z", 36), float("1_000.5"), 1_0.5e1_0 > 0)
print(round(123.456, -1), round(-0.5), round(7, 2), 10 // 3 * 3 + 10 % 3)
nums = [5, 3, 9]
nums[0], nums[-1] = nums[-1], nums[0]
print(nums, nums.index(3))
print(sorted({"b": 1, "a": 2}.items(), key=lambda kv: kv[1], reverse=True))
print(list(zip(range(3), "abc", [True, False, None])))
empty = None
print(empty or "default", empty is None, type(empty).__name__)
print((lambda *a, **k: (a, sorted(k)))(1, 2, z=3, y=4))
print(" ".join(w[::-1] for w in "hello big world".split()))
i = 10
while i:
    i //= 2
print(i, 2 ** 30 - 1, -(2 ** 30))
