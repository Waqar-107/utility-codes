# from dust i have come, dust i will be

import requests
import os


class problem:
    def __init__(self, contestId, index, rating):
        self.contestId = contestId
        self.index = index
        self.rating = rating

    def print_all(self):
        print(self.contestId, self.index, self.rating)


def get_solved_detail(handle, quantity):
    url = "https://codeforces.com/api/user.status?handle=" + handle + "&from=1&count=" + quantity
    res = requests.get(url)

    json_obj = res.json()

    ret = []
    for p in json_obj["result"]:
        if p["verdict"] == "OK" and "contestId" in p:
            # avoid gym problems
            if "rating" in p["problem"]:
                ret.append(problem(p["contestId"], p["problem"]["index"], p["problem"]["rating"]))

    return ret


def get_solved_minimal(handle, quantity):
    url = "https://codeforces.com/api/user.status?handle=" + handle + "&from=1&count=" + quantity
    res = requests.get(url)

    json_obj = res.json()

    ret = []
    for p in json_obj["result"]:
        if p["verdict"] == "OK" and "contestId" in p:
            ret.append(str(p["contestId"]) + p["problem"]["index"])

    return ret


me = "_lucifer_"
target = "CaNDidaTE_FaSTer"

solvedByMe = get_solved_minimal(me, "5000")
solvedByTarget = get_solved_detail(target, "5000")

solvedByTarget.sort(key=lambda pr: (pr.rating, pr.contestId))

# --------------------------------------
# make a folder and put the problems in A,B,C,D,E,F and others file
cwd = os.getcwd()
dir = os.path.join(cwd, "cf_problemset")
if os.path.exists(dir):
    print("directory already exists:", dir)
else:
    os.mkdir(dir)
    print("created directory:", dir)

filePath = os.path.join(dir, "A.txt")
fileA = open(filePath, "w")

filePath = os.path.join(dir, "B.txt")
fileB = open(filePath, "w")

filePath = os.path.join(dir, "C.txt")
fileC = open(filePath, "w")

filePath = os.path.join(dir, "D.txt")
fileD = open(filePath, "w")

filePath = os.path.join(dir, "E.txt")
fileE = open(filePath, "w")

filePath = os.path.join(dir, "F.txt")
fileF = open(filePath, "w")

filePath = os.path.join(dir, "others.txt")
fileOther = open(filePath, "w")
# --------------------------------------

for p in solvedByTarget:
    s = str(p.contestId) + p.index

    if s not in solvedByMe:
        if p.index == "A":
            fileA.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
        elif p.index == "B":
            fileB.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
        elif p.index == "C":
            fileC.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
        elif p.index == "D":
            fileD.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
        elif p.index == "E":
            fileE.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
        elif p.index == "F":
            fileF.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
        else:
            fileOther.write(str(p.contestId) + p.index + " : difficulty => " + str(p.rating) + "\n")
