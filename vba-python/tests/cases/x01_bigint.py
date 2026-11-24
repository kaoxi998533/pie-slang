# Excel only: integers beyond 32 bits use VBA's Decimal type (up to 28 digits),
# which LibreOffice Basic does not implement.
print(2 ** 31, 2 ** 31 - 1, -(2 ** 31), 2 ** 62, 10 ** 20 + 1)
print(sum(range(100000)) * 1000, 123456789 * 987654321)
import math
print(math.factorial(25), 2 ** 64 // 3, 2 ** 64 % 1000, -(10 ** 15) // 7)
print(hex(2 ** 40), int("123456789012345678"), str(9 ** 20)[:5], abs(-(3 ** 30)))
print(float(2 ** 40), 2 ** 40 / 2 ** 20, 2 ** 40 > 2 ** 39, 2 ** 40 == 1099511627776)
d = {2 ** 40: "big"}
print(d[1099511627776], f"{2 ** 40:,}", round(2 ** 35 / 3))
