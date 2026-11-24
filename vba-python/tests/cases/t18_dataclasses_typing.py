from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Tuple

@dataclass
class Point:
    x: float
    y: float = 0.0

    def norm2(self) -> float:
        return self.x ** 2 + self.y ** 2

@dataclass(order=True)
class Version:
    major: int
    minor: int = 0
    tags: List[str] = field(default_factory=list)

p = Point(3, 4)
print(p, p.norm2(), Point(1) == Point(1, 0.0), Point(1) == Point(2), asdict(p))
v1, v2 = Version(1, 2), Version(1, 10, ["beta"])
v1.tags.append("x")
print(v1, v2, v1 < v2, sorted([v2, v1])[0].minor, Version(2).tags)

@dataclass
class Item:
    name: str
    qty: int = 1
    def __post_init__(self):
        self.name = self.name.upper()
print(Item("pen", qty=3), Item(name="ink"))

def total(prices: Dict[str, float], discount: Optional[float] = None) -> float:
    s = sum(prices.values())
    return s * (1 - discount) if discount else s
print(total({"a": 2.0, "b": 3.0}), total({"a": 10.0}, 0.5))
Pair = Tuple[int, int]
def swap(p: Pair) -> Pair:
    return p[1], p[0]
print(swap((1, 2)))
count: int = 5
names: list[str] = ["b", "a"]
print(count, sorted(names))

class Key:
    def __init__(self, a, b):
        self.a, self.b = a, b
    def __eq__(self, o):
        return (self.a, self.b) == (o.a, o.b)
    def __hash__(self):
        return hash((self.a, self.b))
d = {Key(1, 2): "found"}
print(d[Key(1, 2)], Key(1, 2) in d, Key(2, 1) in d, len({Key(1, 1), Key(1, 1)}))
print(hash(7), hash(True), hash(2.0) == hash(2), hash("abc") == hash("abc"))
