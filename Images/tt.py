num = 10
for i in range(10):
    for j in range(2, 15, 3):
        if num % 2 == 0:
            continue
            num += 1
    num += 1
else:
    num += 1
print(num)