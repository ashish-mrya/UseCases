with open("lastUpdated.txt", "w") as lu:
    lu.write("4")

with open("lastUpdated.txt", 'r') as lu:
    last_updated = lu.read()

print(last_updated)
