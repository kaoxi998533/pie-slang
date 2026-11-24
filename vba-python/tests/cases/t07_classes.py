class Animal:
    kingdom = "animalia"
    count = 0
    def __init__(self, name, sound="..."):
        self.name = name
        self.sound = sound
        Animal.count += 1
    def speak(self):
        return f"{self.name} says {self.sound}"
    def __repr__(self):
        return f"{type(self).__name__}({self.name!r})"

class Dog(Animal):
    def __init__(self, name):
        super().__init__(name, "woof")
        self.tricks = []
    def add_trick(self, t):
        self.tricks.append(t)
        return self
    def speak(self):
        return super().speak() + "!"

class Puppy(Dog):
    def speak(self):
        return "small " + super().speak()

a = Animal("cat", "meow")
d = Dog("rex").add_trick("sit").add_trick("roll")
p = Puppy("bit")
print(a.speak(), d.speak(), p.speak(), sep=" | ")
print(d, d.tricks, Animal.count, d.kingdom, isinstance(p, Animal), isinstance(a, Dog))
print(issubclass(Puppy, Animal), type(p) is Puppy, type(p).__name__, [a, d])

class Vector:
    def __init__(self, x, y):
        self.x, self.y = x, y
    def __add__(self, o): return Vector(self.x + o.x, self.y + o.y)
    def __sub__(self, o): return Vector(self.x - o.x, self.y - o.y)
    def __mul__(self, k): return Vector(self.x * k, self.y * k)
    def __rmul__(self, k): return self * k
    def __neg__(self): return Vector(-self.x, -self.y)
    def __eq__(self, o): return isinstance(o, Vector) and (self.x, self.y) == (o.x, o.y)
    def __lt__(self, o): return (self.x ** 2 + self.y ** 2) < (o.x ** 2 + o.y ** 2)
    def __abs__(self): return (self.x ** 2 + self.y ** 2) ** 0.5
    def __len__(self): return 2
    def __getitem__(self, i): return (self.x, self.y)[i]
    def __iter__(self): return iter([self.x, self.y])
    def __str__(self): return f"<{self.x}, {self.y}>"
    def __repr__(self): return f"Vector({self.x}, {self.y})"
    def __hash__(self): return hash((self.x, self.y))
    def __bool__(self): return bool(self.x or self.y)
    def __contains__(self, v): return v in (self.x, self.y)

v, w = Vector(1, 2), Vector(3, 4)
print(v + w, v - w, v * 3, 2 * v, -v, abs(w), len(v), v[0], v[1], list(v))
print(v == Vector(1, 2), v != w, v < w, sorted([w, v]), max([v, w]), str(v), repr(v))
print(bool(Vector(0, 0)), bool(v), 2 in v, 5 in v)
x, y = w
print(x, y)

class Temperature:
    def __init__(self, c):
        self._c = c
    @property
    def celsius(self):
        return self._c
    @celsius.setter
    def celsius(self, value):
        if value < -273.15:
            raise ValueError("too cold")
        self._c = value
    @property
    def fahrenheit(self):
        return self._c * 9 / 5 + 32
    @staticmethod
    def describe():
        return "temperature"
    @classmethod
    def freezing(cls):
        return cls(0)

t = Temperature(25)
t.celsius = 30
print(t.celsius, t.fahrenheit, Temperature.describe(), Temperature.freezing().celsius, t.describe())
try:
    t.celsius = -300
except ValueError as e:
    print("error:", e)

class Counter:
    def __init__(self, limit):
        self.limit, self.n = limit, 0
    def __iter__(self):
        return self
    def __next__(self):
        if self.n >= self.limit:
            raise StopIteration
        self.n += 1
        return self.n

print(list(Counter(4)), sum(Counter(10)), [i * i for i in Counter(3)])
it = iter([1, 2])
print(next(it), next(it), next(it, "done"))

class Adder:
    def __init__(self, n): self.n = n
    def __call__(self, x): return x + self.n
add5 = Adder(5)
print(add5(10), callable(add5), callable(3), list(map(add5, [1, 2])))

class Node:
    def __init__(self, val, next=None):
        self.val, self.next = val, next
def to_list(node):
    out = []
    while node:
        out.append(node.val)
        node = node.next
    return out
print(to_list(Node(1, Node(2, Node(3)))))
class Empty: pass
e = Empty()
e.x = 5
setattr(e, "y", 6)
print(e.x, getattr(e, "y"), getattr(e, "z", "none"), hasattr(e, "x"), hasattr(e, "q"))
print(sorted(vars(e).items()))
class Base:
    def hello(self): return "base " + self.name()
    def name(self): return "B"
class Mixin:
    def extra(self): return "mixin"
class Child(Base, Mixin):
    def name(self): return "C"
c = Child()
print(c.hello(), c.extra(), isinstance(c, Mixin))
