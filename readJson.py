import json

#opening JSON file
with open("config.json") as f:
    data = json.load(f)
#retrun JSON object as a dictionary

print(data)




#iterate over the Json list