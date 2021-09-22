def meraki_helper(n):
    digits = []
    while n:
            digits.append(n % 10)
            n -= n % 10
            n /= 10
    for i in range(1, len(digits)):
            if abs(digits[i] - digits[i - 1]) != 1:
                return 0
    return 1
             input_list = [12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321]
    for i in input_list:
             z = meraki_helper(i)
        if z:
            print("Yes - " + str(i) + " is a Meraki number!!")
        else:
            print("No - " + str(i) + " is not a Meraki number!!")