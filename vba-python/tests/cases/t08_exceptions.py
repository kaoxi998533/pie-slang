def div(a, b):
    try:
        r = a / b
    except ZeroDivisionError as e:
        print("caught:", e)
        return None
    else:
        print("no error")
        return r
    finally:
        print("finally", a, b)

print(div(1, 2))
print(div(1, 0))

class AppError(Exception):
    pass
class NotFound(AppError):
    def __init__(self, what):
        super().__init__(f"{what} not found")
        self.what = what

def find(key):
    data = {"a": 1}
    if key not in data:
        raise NotFound(key)
    return data[key]

for k in ["a", "b"]:
    try:
        print("found", find(k))
    except AppError as e:
        print(type(e).__name__, e, e.what, e.args)

try:
    [1, 2][5]
except IndexError as e:
    print("IndexError:", e)
try:
    {}["k"]
except KeyError as e:
    print("KeyError:", e)
try:
    int("abc")
except ValueError as e:
    print("ValueError:", e)
try:
    None + 1
except TypeError as e:
    print("TypeError:", e)
try:
    undefined_name
except NameError as e:
    print("NameError:", e)
try:
    "x".nope
except AttributeError as e:
    print("AttributeError:", e)
try:
    raise ValueError("bad", 42)
except (TypeError, ValueError) as e:
    print("tuple catch", e.args)

def reraise():
    try:
        1 / 0
    except ZeroDivisionError:
        print("logging")
        raise
try:
    reraise()
except Exception as e:
    print("outer got", repr(e))

def nested():
    try:
        try:
            raise KeyError("inner")
        finally:
            print("inner finally")
    except KeyError as e:
        print("outer except", e)
        return "returned"
    finally:
        print("outer finally")
print(nested())

def finally_return():
    try:
        return "try"
    finally:
        print("cleanup")
print(finally_return())

count = 0
while True:
    try:
        count += 1
        if count < 3:
            raise RuntimeError("again")
        break
    except RuntimeError:
        continue
print("count", count)

try:
    assert 1 + 1 == 3, "math is broken"
except AssertionError as e:
    print("assert:", e)
e = Exception("plain")
print(str(e), repr(e), e.args)
try:
    raise NotImplementedError
except Exception as ex:
    print(type(ex).__name__, isinstance(ex, RuntimeError))
def deep(n):
    if n == 0:
        raise ValueError("bottom")
    return deep(n - 1)
try:
    deep(50)
except ValueError as e:
    print("deep:", e)
try:
    try:
        1 / 0
    except ZeroDivisionError:
        raise ValueError("converted")
except ValueError as e:
    print(e)
