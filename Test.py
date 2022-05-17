#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Есть словарь координат городов

cites = {
    'Moscow': (550, 370),
    'London': (510, 510),
    'Paris': (480, 480),
}

# Составим словарь словарей расстояний между ними
# расстояние на координатной сетке - корень из (x1 - x2) ** 2 + (y1 - y2) ** 2

distances = {}

mos = cites['Moscow']
lond = cites["London"]
paris = cites ["Paris"]

mos_lond = ((mos[0]-lond[0])**2+(mos[1]-lond[1])**2)**0.5
mos_paris = ((mos[0]-paris[0])**2+(mos[1]-paris[1])**2)**0.5
lond_paris = ((paris[0]-lond[0])**2+(paris[1]-lond[1])**2)**0.5

distances ["Moscow"] = {}
distances ["Moscow"]["London"] = mos_lond
distances ["Moscow"]["Paris"] = mos_paris

print(distances)
print(distances["Moscow"]["Paris"])






