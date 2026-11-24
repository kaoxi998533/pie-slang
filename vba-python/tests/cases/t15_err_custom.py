class InsufficientFunds(Exception):
    pass
def withdraw(balance, amt):
    if amt > balance:
        raise InsufficientFunds(f"need {amt - balance} more")
    return balance - amt
print(withdraw(10, 3))
withdraw(10, 30)
