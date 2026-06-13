#!/usr/bin/env python3
"""
Produce a skill rating matrix: rows = user_id, columns = game number (1, 2, 3, …).
Cell value = new_skill - new_uncertainty. Rows in match_players.csv are already
sorted ascending per player, so no sort/merge needed.

Usage:
    python skill_matrix.py <match_players.csv> [output.xlsx]
"""

import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule


def main():
    output_path = sys.argv[1] if len(sys.argv) > 1 else "skill_matrix.csv"

    players = {}
    matches = set()

    with open("matches.csv") as file:
        i=0
        for line in file:
            if i>0:
                fields = line.rstrip().split(',')
                matchid = fields[0]
                game_type = fields[6]
                if game_type == "Large Team":
                    matches.add(str(matchid))
            i+=1


    max = 0
    with open("match_players.csv") as file:
        i=0
        for line in file:
            if i>0:
                fields = line.rstrip().split(',')
                matchid = fields[0]
                playerid = fields[2]
                if str(matchid) not in matches:
                    continue
                if len(fields[7])>0 and len(fields[8])>0:
                    skill = float(fields[7]) - float(fields[8])
                    if playerid not in players:
                        players[playerid] = []
                    players[playerid].append(skill)
                    if len(players[playerid]) > max:
                        max = len(players[playerid])
            i+=1

    print("Found", len(players), "players")

    with open(output_path, 'w') as the_file:
        the_file.write('user_id')
        for i in range(0, max):
            the_file.write(",")
            the_file.write(str(i))
        the_file.write("\n")
        for index, (k, vs) in enumerate(players.items()):
            #print("k", k)
            #print("vs", vs)
            the_file.write(str(k))
            for v in vs:
                the_file.write(",")
                the_file.write(str(v))
            the_file.write("\n")

if __name__ == "__main__":
    main()