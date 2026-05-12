n = input("SPT Blow count:").split(" ")
print(f"\nAnalyzing {len(n)} data points...\n")
for i in range(len(n)):
    n[i] = float(n[i])
    n60 = n[i] * 0.75
    if n60 < 4:
        desc = "Very Loose"
    elif 4 < n60 < 10:
        desc = "Loose"
    elif 10 < n60 < 30:
        desc = "Medium Dense"
    elif 30 < n60 < 50:
        desc = "Dense"
    elif n60 >= 50:
        desc = "Very Dense"
    print(f"SPT Blow count: {n[i]:<4}, N60: {n60:<10}, Description: {desc}")
