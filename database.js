// Chứa toàn bộ dữ liệu Excel thô 
const rawData = `
| Image | Serial Code | Pokémon | Rarity | Type 1 | Type 2 | Gimmick | Base PE | Gimmick PE | Level | Ability | HP | Attack | Mega Atk | Defense | Mega Def | Sp. Attack | Mega Sp.A | Sp. Defense | Mega Sp.D | Speed | Mega Spe | Move 1 | Type 1 (M1) | Category 1 |
| ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ | ------ |
|  | 1-1-001 | Mewtwo | 6-Star | Psychic |  |  | 158 |  | 52 |  | 172 | 119 |  | 98 |  | 165 |  | 98 |  | 140 |  | Psystrike | Psychic | Special |
|  | 1-1-002 | Mew | 6-Star | Psychic |  |  | 142 |  | 52 |  | 166 | 109 |  | 109 |  | 109 |  | 109 |  | 109 |  | Psychic | Psychic | Special |
|  | 1-1-003 | Zacian Crowned | 6-Star | Fairy | Steel |  | 152 |  | 50 |  | 152 | 175 |  | 120 |  | 85 |  | 120 |  | 153 |  | Behemoth Blade | Steel | Physical |
|  | 1-1-004 | Zamazenta Crowned | 6-Star | Fighting | Steel |  | 152 |  | 50 |  | 152 | 135 |  | 150 |  | 85 |  | 150 |  | 133 |  | Behemoth Bash | Steel | Physical |
|  | 1-1-005 | Tyranitar | 6-Star | Rock | Dark | Dynamax | 144 | 152/159/166 | 53 |  | 169 | 147 |  | 121 |  | 105 |  | 111 |  | 69 |  | Stone Edge | Rock | Physical |
|  | 1-1-006 | Metagross | 6-Star | Steel | Psychic | Dynamax | 144 | 152/159/166 | 53 |  | 147 | 148 |  | 142 |  | 105 |  | 100 |  | 79 |  | Meteor Mash | Steel | Physical |
|  | 1-1-007 | Miraidon | 6-Star | Electric | Dragon |  | 162 |  | 54 |  | 172 | 96 |  | 113 |  | 150 |  | 129 |  | 150 |  | Electro Drift | Electric | Special |
|  | 1-1-008 | Venusaur Mega | 6-Star | Grass | Poison | Mega Evolution | 122 | 126 | 50 |  | 140 | 87 | 105 | 88 | 128 | 105 | 127 | 105 | 125 | 85 |  | Leaf Storm | Grass | Special |
|  | 1-1-009 | Charizard Mega X | 6-Star | Fire | Flying | Mega Evolution | 124 | 124 | 50 |  | 138 | 89 | 135 | 83 | 110 | 114 | 135 | 90 |  | 105 |  | Flamethrower | Fire | Special |
|  | 1-1-010 | Blastoise Mega | 6-Star | Water |  | Mega Evolution | 122 | 126 | 50 |  | 139 | 88 | 108 | 105 | 125 | 90 | 140 | 110 | 120 | 83 |  | Hydro Pump | Water | Special |
|  | 1-1-011 | Vaporeon | 5-Star | Water |  |  | 108 |  | 44 |  | 168 | 62 |  | 57 |  | 101 |  | 88 |  | 62 |  | Hydro Pump | Water | Special |
|  | 1-1-012 | Flareon | 5-Star | Fire |  |  | 108 |  | 44 |  | 111 | 119 |  | 57 |  | 88 |  | 101 |  | 62 |  | Flamethrower | Fire | Special |
|  | 1-1-013 | Umbreon | 5-Star | Dark |  |  | 112 |  | 46 |  | 143 | 64 |  | 106 |  | 60 |  | 124 |  | 64 |  | Dark Pulse | Dark | Special |
|  | 1-1-014 | Glaceon | 5-Star | Ice |  |  | 114 |  | 47 |  | 118 | 61 |  | 108 |  | 127 |  | 94 |  | 66 |  | Blizzard | Ice | Special |
|  | 1-1-015 | Pidgeot Mega | 5-Star | Normal | Flying | Mega Evolution | 100 | 102 | 44 |  | 127 | 75 |  | 71 | 75 | 66 | 123 | 66 | 75 | 93 | 111 | Hurricane | Flying | Special |
|  | 1-1-016 | Gengar Mega | 5-Star | Ghost | Poison | Mega Evolution | 108 | 114 | 46 |  | 111 | 64 |  | 60 | 78 | 124 | 101 | 74 | 92 | 106 | 124 | Shadow Ball | Ghost | Special |
|  | 1-1-017 | Arcanine | 5-Star | Fire |  | Z-Move | 114 |  | 44 |  | 133 | 101 |  | 75 |  | 93 |  | 75 |  | 88 |  | Fire Blast | Fire | Special |
|  | 1-1-018 | Rhyperior | 5-Star | Ground | Rock | Z-Move | 110 |  | 44 |  | 155 | 128 |  | 119 |  | 53 |  | 53 |  | 39 |  | Earthquake | Ground | Physical |
|  | 1-1-019 | Pikachu | 5-Star | Electric |  | Dynamax | 90 | 101/106/111 | 51 |  | 113 | 72 |  | 53 |  | 66 |  | 66 |  | 114 |  | Thunderbolt | Electric | Special |
|  | 1-1-020 | Flygon | 5-Star | Ground | Dragon | Dynamax | 114 | 111 / 120 / 125 / 132 | 47 |  | 132 | 99 |  | 80 |  | 80 |  | 80 |  | 99 |  | Earthquake | Ground | Physical |
|  | 1-1-021 | Rillaboom | 5-Star | Grass |  |  | 109 |  | 44 |  | 142 | 115 |  | 84 |  | 57 |  | 66 |  | 79 |  | Drum Beating | Grass | Physical |
|  | 1-1-022 | Cinderace | 5-Star | Fire |  |  | 109 |  | 44 |  | 124 | 107 |  | 71 |  | 62 |  | 71 |  | 109 |  | Pyro Ball | Fire | Physical |
|  | 1-1-023 | Inteleon | 5-Star | Water |  |  | 109 |  | 44 |  | 115 | 79 |  | 62 |  | 115 |  | 62 |  | 110 |  | Snipe Shot | Water | Special |
|  | 1-1-024 | Obstagoon | 5-Star | Dark | Normal |  | 112 |  | 45 |  | 141 | 87 |  | 97 |  | 60 |  | 79 |  | 92 |  | Throat Chop | Dark | Physical |
|  | 1-1-025 | Keldeo | 5-Star | Water | Fighting |  | 122 |  | 46 |  | 139 | 71 |  | 87 |  | 123 |  | 87 |  | 104 |  | Sacred Sword | Fighting | Physical |
|  | 1-1-026 | Sprigatito | 4-Star | Grass |  |  | 60 |  | 24 |  | 60 | 39 |  | 35 |  | 30 |  | 30 |  | 41 |  | Leafage | Grass | Physical |
|  | 1-1-027 | Floragato | 4-Star | Grass |  |  | 84 |  | 36 |  | 103 | 72 |  | 57 |  | 55 |  | 57 |  | 74 |  | Magical Leaf | Grass | Special |
|  | 1-1-028 | Fuecoco | 4-Star | Fire |  |  | 60 |  | 24 |  | 75 | 30 |  | 38 |  | 40 |  | 27 |  | 25 |  | Ember | Fire | Special |
|  | 1-1-029 | Crocalor | 4-Star | Fire |  |  | 84 |  | 36 |  | 120 | 51 |  | 70 |  | 80 |  | 53 |  | 46 |  | Incinerate | Fire | Special |
|  | 1-1-030 | Quaxly | 4-Star | Water |  |  | 60 |  | 24 |  | 68 | 41 |  | 30 |  | 33 |  | 30 |  | 33 |  | Water Gun | Water | Special |
|  | 1-1-031 | Quaxwell | 4-Star | Water |  |  | 84 |  | 36 |  | 110 | 76 |  | 59 |  | 59 |  | 55 |  | 59 |  | Water Pulse | Water | Special |
|  | 1-1-032 | Grookey | 4-Star | Grass |  |  | 58 |  | 24 |  | 58 | 36 |  | 29 |  | 24 |  | 24 |  | 30 |  | Branch Poke | Grass | Physical |
|  | 1-1-033 | Thwackey | 4-Star | Grass |  |  | 82 |  | 36 |  | 82 | 56 |  | 47 |  | 38 |  | 41 |  | 53 |  | Razor Leaf | Grass | Physical |
|  | 1-1-034 | Rillaboom | 4-Star | Grass |  |  | 98 |  | 36 |  | 118 | 95 |  | 69 |  | 48 |  | 55 |  | 68 |  | Bullet Seed | Grass | Physical |
|  | 1-1-035 | Scorbunny | 4-Star | Fire |  |  | 58 |  | 24 |  | 58 | 40 |  | 24 |  | 24 |  | 24 |  | 38 |  | Ember | Fire | Special |
|  | 1-1-036 | Raboot | 4-Star | Fire |  |  | 79 |  | 36 |  | 79 | 56 |  | 41 |  | 38 |  | 41 |  | 61 |  | Flame Charge | Fire | Physical |
|  | 1-1-037 | Cinderace | 4-Star | Fire |  |  | 98 |  | 36 |  | 103 | 88 |  | 59 |  | 51 |  | 59 |  | 90 |  | Blaze Kick | Fire | Physical |
|  | 1-1-038 | Sobble | 4-Star | Water |  |  | 58 |  | 24 |  | 58 | 24 |  | 24 |  | 39 |  | 24 |  | 39 |  | Water Gun | Water | Special |
|  | 1-1-039 | Drizzile | 4-Star | Water |  |  | 84 |  | 36 |  | 79 | 41 |  | 38 |  | 62 |  | 38 |  | 59 |  | Water Pulse | Water | Special |
|  | 1-1-040 | Inteleon | 4-Star | Water |  |  | 98 |  | 36 |  | 96 | 60 |  | 51 |  | 95 |  | 51 |  | 91 |  | Dive | Water | Physical |
|  | 1-1-041 | Dedenne | 4-Star | Electric | Fairy |  | 52 |  | 24 |  | 66 | 32 |  | 32 |  | 44 |  | 38 |  | 55 |  | Thunder Shock | Electric | Special |
|  | 1-1-042 | Pikachu | 4-Star | Electric |  |  | 60 |  | 30 |  | 71 | 44 |  | 33 |  | 41 |  | 41 |  | 70 |  | Spark | Electric | Physical |
|  | 1-1-043 | Raichu | 4-Star | Electric |  |  | 84 |  | 36 |  | 89 | 69 |  | 44 |  | 69 |  | 62 |  | 84 |  | Thunder | Electric | Special |
|  | 1-1-044 | Eevee | 4-Star | Normal |  |  | 60 |  | 30 |  | 85 | 44 |  | 41 |  | 37 |  | 51 |  | 44 |  | Swift | Normal | Special |
|  | 1-1-045 | Jolteon | 4-Star | Electric |  |  | 92 |  | 37 |  | 92 | 51 |  | 48 |  | 84 |  | 73 |  | 99 |  | Thunder | Electric | Special |
|  | 1-1-046 | Espeon | 4-Star | Psychic |  |  | 92 |  | 37 |  | 95 | 53 |  | 49 |  | 101 |  | 75 |  | 86 |  | Psybeam | Psychic | Special |
|  | 1-1-047 | Leafeon | 4-Star | Grass |  |  | 94 |  | 38 |  | 97 | 88 |  | 103 |  | 50 |  | 54 |  | 77 |  | Razor Leaf | Grass | Physical |
|  | 1-1-048 | Sylveon | 4-Star | Fairy |  |  | 96 |  | 39 |  | 123 | 55 |  | 55 |  | 90 |  | 106 |  | 51 |  | Dazzling Gleam | Fairy | Special |
|  | 1-1-049 | Sneasel | 3-Star | Dark | Ice |  | 68 |  | 31 |  | 75 | 63 |  | 39 |  | 26 |  | 51 |  | 76 |  | Payback | Dark | Physical |
|  | 1-1-050 | Weavile | 4-Star | Dark | Ice |  | 92 |  | 37 |  | 98 | 93 |  | 53 |  | 37 |  | 67 |  | 97 |  | Night Slash | Dark | Physical |
|  | 1-1-051 | Golett | 4-Star | Ground | Ghost |  | 62 |  | 25 |  | 64 | 42 |  | 30 |  | 22 |  | 30 |  | 22 |  | Astonish | Ghost | Physical |
|  | 1-1-052 | Golurk | 4-Star | Ground | Ghost |  | 72 |  | 31 |  | 96 | 81 |  | 54 |  | 39 |  | 54 |  | 39 |  | Shadow Punch | Ghost | Physical |
|  | 1-1-053 | Slowpoke | 4-Star | Water | Psychic |  | 62 |  | 24 |  | 77 | 37 |  | 37 |  | 24 |  | 24 |  | 9 |  | Water Gun | Water | Special |
|  | 1-1-054 | Slowbro | 4-Star | Water | Psychic |  | 84 |  | 36 |  | 114 | 59 |  | 84 |  | 77 |  | 62 |  | 24 |  | Water Pulse | Water | Special |
|  | 1-1-055 | Mudbray | 3-Star | Ground |  |  | 62 |  | 25 |  | 86 | 69 |  | 49 |  | 33 |  | 40 |  | 33 |  | Bulldoze | Ground | Physical |
|  | 1-1-056 | Mudsdale | 4-Star | Ground |  |  | 90 |  | 38 |  | 124 | 100 |  | 81 |  | 46 |  | 69 |  | 31 |  | Stomping Tantrum | Ground | Physical |
|  | 1-1-057 | Zigzagoon (Galarian) | 4-Star | Dark | Normal |  | 36 |  | 25 |  | 54 | 20 |  | 25 |  | 20 |  | 25 |  | 35 |  | Payback | Dark | Physical |
|  | 1-1-058 | Linoone (Galarian) | 4-Star | Dark | Normal |  | 64 |  | 31 |  | 89 | 48 |  | 42 |  | 36 |  | 42 |  | 67 |  | Snarl | Dark | Special |
|  | 1-1-059 | Obstagoon | 4-Star | Dark | Normal |  | 90 |  | 37 |  | 115 | 71 |  | 79 |  | 49 |  | 64 |  | 75 |  | Assurance | Dark | Physical |
|  | 1-1-060 | Grubbin | 4-Star | Bug |  |  | 42 |  | 25 |  | 58 | 36 |  | 27 |  | 32 |  | 27 |  | 28 |  | Bug Bite | Bug | Physical |
|  | 1-1-061 | Charjabug | 4-Star | Bug | Electric |  | 62 |  | 31 |  | 76 | 55 |  | 63 |  | 39 |  | 51 |  | 26 |  | X-Scissor | Bug | Physical |
|  | 1-1-062 | Vikavolt | 4-Star | Bug | Electric |  | 88 |  | 37 |  | 103 | 56 |  | 71 |  | 113 |  | 60 |  | 35 |  | Bug Buzz | Bug | Special |
|  | 1-1-063 | Doduo | 4-Star | Normal | Flying |  | 40 |  | 24 |  | 50 | 46 |  | 26 |  | 21 |  | 21 |  | 41 |  | Peck | Flying | Physical |
|  | 1-1-064 | Dodrio | 4-Star | Normal | Flying |  | 82 |  | 35 |  | 89 | 84 |  | 55 |  | 48 |  | 48 |  | 84 |  | Pluck | Flying | Physical |
|  | 1-1-065 | Gossifleur | 4-Star | Grass |  |  | 44 |  | 30 |  | 64 | 28 |  | 41 |  | 28 |  | 41 |  | 7 |  | Leafage | Grass | Physical |
|  | 1-1-066 | Eldegoss | 4-Star | Grass |  |  | 80 |  | 35 |  | 89 | 41 |  | 69 |  | 62 |  | 92 |  | 48 |  | Leaf Tornado | Grass | Special |
|  | 1-1-067 | Turtonator | 4-Star | Fire | Dragon |  | 84 |  | 36 |  | 89 | 61 |  | 102 |  | 70 |  | 66 |  | 30 |  | Incinerate | Fire | Special |
|  | 1-1-068 | Cramorant | 3-Star | Flying | Water |  | 70 |  | 30 |  | 82 | 56 |  | 38 |  | 56 |  | 62 |  | 56 |  | Scald | Water | Special |
|  | 1-1-069 | Snom | 2-Star | Ice | Bug |  | 38 |  | 25 |  | 50 | 17 |  | 22 |  | 29 |  | 20 |  | 14 |  | Powder Snow | Ice | Special |
|  | 1-1-070 | Frosmoth | 4-Star | Ice | Bug |  | 88 |  | 35 |  | 101 | 54 |  | 50 |  | 100 |  | 73 |  | 54 |  | Aurora Beam | Ice | Special |
|  | 1-2-001 | Kyogre | 6-Star | Water |  |  | 152 |  | 50 |  | 160 | 105 |  | 95 |  | 155 |  | 145 |  | 95 |  | Hydro Pump | Water | Special |
|  | 1-2-002 | Groudon | 6-Star | Ground |  |  | 152 |  | 50 |  | 160 | 155 |  | 145 |  | 105 |  | 95 |  | 95 |  | Earthquake | Ground | Physical |
|  | 1-2-003 | Koraidon | 6-Star | Fighting | Dragon |  | 162 |  | 54 |  | 172 | 150 |  | 129 |  | 96 |  | 113 |  | 150 |  | Collision Course | Fighting | Physical |
|  | 1-2-004 | Pikachu | 6-Star | Electric |  | Z-Move | 102 |  | 44 |  | 118 | 75 |  | 56 |  | 69 |  | 69 |  | 120 |  | Thunderbolt | Electric | Special |
|  | 1-2-005 | Snorlax | 6-Star | Normal |  | Dynamax | 144 | 152 / 159 / 166 | 50 |  | 253 | 132 |  | 80 |  | 80 |  | 132 |  | 39 |  | Giga Impact | Normal | Physical |
|  | 1-2-006 | Kommo-o | 6-Star | Dragon | Fighting | Dynamax | 150 | 159 / 166 / 174 / 182 | 50 |  | 155 | 132 |  | 150 |  | 121 |  | 126 |  | 103 |  | Focus Blast | Fighting | Special |
|  | 1-2-007 | Gardevoir (Mega) | 6-Star | Psychic | Fairy | Mega Evolution | 124 | 132 | 52 |  | 132 | 72 | 93 | 72 |  | 135 | 170 | 124 | 145 | 88 | 109 | Psychic | Psychic | Special |
|  | 1-2-008 | Reshiram | 6-Star | Dragon | Fire |  | 162 |  | 53 |  | 169 | 132 |  | 111 |  | 164 |  | 132 |  | 100 |  | Fusion Flare | Fire | Special |
|  | 1-2-009 | Zekrom | 6-Star | Dragon | Electric |  | 162 |  | 53 |  | 169 | 164 |  | 132 |  | 132 |  | 111 |  | 100 |  | Fusion Bolt | Electric | Physical |
|  | 1-2-010 | Kyurem | 6-Star | Dragon | Ice |  | 150 |  | 53 |  | 195 | 142 |  | 100 |  | 142 |  | 100 |  | 105 |  | Blizzard | Ice | Special |
|  | 1-2-011 | Jolteon | 5-Star | Electric |  |  | 100 |  | 44 |  | 111 | 62 |  | 57 |  | 101 |  | 88 |  | 119 |  | Thunderbolt | Electric | Special |
|  | 1-2-012 | Espeon | 5-Star | Psychic |  |  | 112 |  | 46 |  | 115 | 64 |  | 60 |  | 124 |  | 92 |  | 106 |  | Psychic | Psychic | Special |
|  | 1-2-013 | Leafeon | 5-Star | Grass |  |  | 114 |  | 47 |  | 118 | 108 |  | 127 |  | 61 |  | 66 |  | 94 |  | Leaf Blade | Grass | Physical |
|  | 1-2-014 | Sylveon | 5-Star | Fairy |  |  | 118 |  | 48 |  | 149 | 67 |  | 67 |  | 110 |  | 129 |  | 62 |  | Moonblast | Fairy | Special |
|  | 1-2-015 | Lucario (Mega) | 5-Star | Fighting | Steel | Mega Evolution | 110 | 122 | 49 |  | 125 | 110 | 144 | 72 | 89 | 115 | 139 | 72 |  | 91 | 112 | Aura Sphere | Fighting | Special |
|  | 1-2-016 | Torterra | 5-Star | Grass | Ground | Z-Move | 112 |  | 48 |  | 143 | 105 |  | 101 |  | 74 |  | 83 |  | 56 |  | Earthquake | Ground | Physical |
|  | 1-2-017 | Infernape | 5-Star | Fire | Fighting | Z-Move | 114 |  | 48 |  | 125 | 100 |  | 70 |  | 100 |  | 70 |  | 104 |  | Close Combat | Fighting | Physical |
|  | 1-2-018 | Empoleon | 5-Star | Water | Steel | Z-Move | 114 |  | 48 |  | 133 | 84 |  | 85 |  | 107 |  | 97 |  | 60 |  | Flash Cannon | Steel | Special |
|  | 1-2-019 | Ninetales(Alolan) | 5-Star | Ice | Fairy | Z-Move | 122 |  | 52 |  | 137 | 74 |  | 83 |  | 89 |  | 109 |  | 118 |  | Blizzard | Ice | Special |
|  | 1-2-020 | Tyranitar | 5-Star | Rock | Dark | Dynamax | 132 | 139 / 146 / 152 | 48 |  | 154 | 133 |  | 110 |  | 96 |  | 101 |  | 63 |  | Stone Edge | Rock | Physical |
|  | 1-2-021 | Metagross | 5-Star | Steel | Psychic | Dynamax | 132 | 139 / 146 / 152 | 48 |  | 134 | 134 |  | 129 |  | 96 |  | 91 |  | 72 |  | Meteor Mash | Steel | Physical |
|  | 1-2-022 | Flygon | 5-Star | Grass | Dragon |  | 110 |  | 46 |  | 125 | 110 |  | 81 |  | 96 |  | 62 |  | 72 |  | Dragon Rush | Dragon | Physical |
|  | 1-2-023 | Appletun | 5-Star | Grass | Dragon |  | 110 |  | 40 |  | 163 | 86 |  | 81 |  | 101 |  | 81 |  | 33 |  | Apple Acid | Grass | Special |
|  | 1-2-024 | Alcremie | 5-Star | Fairy |  |  | 112 |  | 40 |  | 120 | 62 |  | 77 |  | 110 |  | 121 |  | 66 |  | Dazzling Gleam | Fairy | Special |
|  | 1-2-025 | Drednaw | 5-Star | Water | Rock |  | 104 |  | 40 |  | 138 | 110 |  | 87 |  | 48 |  | 67 |  | 73 |  | Head Smash | Rock | Physical |
|  | 1-2-026 | Meowscarada | 4-Star | Grass | Dark |  | 92 |  | 37 |  | 103 | 86 |  | 56 |  | 64 |  | 56 |  | 96 |  | Night Slash | Dark | Physical |
|  | 1-2-027 | Skeledirge | 4-Star | Fire | Ghost |  | 92 |  | 37 |  | 123 | 60 |  | 79 |  | 86 |  | 60 |  | 53 |  | Hex | Ghost | Special |
|  | 1-2-028 | Quaquaval | 4-Star | Water | Fighting |  | 92 |  | 37 |  | 109 | 93 |  | 64 |  | 67 |  | 60 |  | 67 |  | Low Sweep | Fighting | Physical |
|  | 1-2-029 | Charcadet | 2-Star | Fire |  |  | 40 |  | 24 |  | 60 | 33 |  | 27 |  | 33 |  | 27 |  | 24 |  | Ember | Fire | Special |
|  | 1-2-030 | Frigibax | 2-Star | Dragon | Ice |  | 48 |  | 24 |  | 74 | 47 |  | 30 |  | 24 |  | 30 |  | 35 |  | Tackle | Normal | Physical |
|  | 1-2-031 | Pawmi | 2-Star | Electric |  |  | 40 |  | 24 |  | 63 | 33 |  | 16 |  | 27 |  | 19 |  | 38 |  | Scratch | Normal | Physical |
|  | 1-2-032 | Turtwig | 2-Star | Grass |  |  | 44 |  | 25 |  | 62 | 39 |  | 37 |  | 27 |  | 32 |  | 20 |  | Absorb | Grass | Special |
|  | 1-2-033 | Grotle | 3-Star | Grass |  |  | 64 |  | 31 |  | 87 | 60 |  | 57 |  | 39 |  | 45 |  | 26 |  | Mega Drain | Grass | Special |
|  | 1-2-034 | Torterra | 4-Star | Grass | Ground |  | 92 |  | 37 |  | 117 | 85 |  | 82 |  | 60 |  | 67 |  | 46 |  | Stomping Tantrum | Ground | Physical |
|  | 1-2-035 | Chimchar | 2-Star | Fire |  |  | 42 |  | 25 |  | 57 | 34 |  | 27 |  | 34 |  | 27 |  | 35 |  | Ember | Fire | Special |
|  | 1-2-036 | Monferno | 3-Star | Fire | Fighting |  | 64 |  | 31 |  | 80 | 53 |  | 37 |  | 53 |  | 37 |  | 55 |  | Mach Punch | Fighting | Physical |
|  | 1-2-037 | Infernape | 4-Star | Fire | Fighting |  | 92 |  | 37 |  | 103 | 81 |  | 57 |  | 81 |  | 57 |  | 84 |  | Brick Break | Fighting | Physical |
|  | 1-2-038 | Piplup | 2-Star | Water |  |  | 44 |  | 25 |  | 61 | 30 |  | 31 |  | 35 |  | 33 |  | 25 |  | Bubble | Water | Special |
|  | 1-2-039 | Prinplup | 3-Star | Water |  |  | 64 |  | 31 |  | 80 | 45 |  | 47 |  | 55 |  | 52 |  | 36 |  | Bubble Beam | Water | Special |
|  | 1-2-040 | Empoleon | 4-Star | Water | Steel |  | 92 |  | 37 |  | 109 | 68 |  | 70 |  | 87 |  | 79 |  | 49 |  | Steel Wing | Steel | Physical |
|  | 1-2-041 | Vaporeon | 4-Star | Water |  |  | 90 |  | 36 |  | 139 | 51 |  | 48 |  | 84 |  | 73 |  | 51 |  | Muddy Water | Water | Special |
|  | 1-2-042 | Flareon | 4-Star | Fire |  |  | 90 |  | 36 |  | 92 | 99 |  | 48 |  | 73 |  | 84 |  | 51 |  | Lava Plume | Fire | Special |
|  | 1-2-043 | Umbreon | 4-Star | Dark |  |  | 92 |  | 37 |  | 117 | 53 |  | 86 |  | 49 |  | 101 |  | 53 |  | Bite | Dark | Physical |
|  | 1-2-044 | Glaceon | 4-Star | Ice |  |  | 94 |  | 38 |  | 97 | 50 |  | 88 |  | 103 |  | 77 |  | 54 |  | Freeze-Dry | Ice | Special |
|  | 1-2-045 | Yamper |  | Electric |  |  | 38 |  | 24 |  | 62 | 26 |  | 29 |  | 24 |  | 29 |  | 16 |  | Nuzzle | Electric | Physical |
|  | 1-2-046 | Boltund | 4-Star | Electric |  |  | 84 |  | 36 |  | 95 | 69 |  | 48 |  | 69 |  | 48 |  | 93 |  | Spark | Electric | Physical |
|  | 1-2-047 | Pidgey |  | Normal | Flying |  | 38 |  | 24 |  | 53 | 26 |  | 24 |  | 21 |  | 21 |  | 31 |  | Gust | Flying | Special |
|  | 1-2-048 | Pidgeotto |  | Normal | Flying |  | 64 |  | 30 |  | 77 | 41 |  | 38 |  | 35 |  | 35 |  | 47 |  | Wing Attack | Flying | Physical |
|  | 1-2-049 | Pidgeot | 4-Star | Normal | Flying |  | 82 |  | 36 |  | 105 | 62 |  | 59 |  | 55 |  | 55 |  | 77 |  | Air Slash | Flying | Special |
|  | 1-2-050 | Sandshrew(Alolan) |  | Ice | Steel |  | 40 |  | 24 |  | 58 | 41 |  | 50 |  | 6 |  | 21 |  | 24 |  | Icy Wind | Ice | Special |
|  | 1-2-051 | Sandslash(Alolan) | 4-Star | Ice | Steel |  | 78 |  | 36 |  | 100 | 77 |  | 92 |  | 20 |  | 51 |  | 51 |  | Icicle Crash | Ice | Physical |
|  | 1-2-052 | Bronzor |  | Steel | Psychic |  | 42 |  | 25 |  | 63 | 17 |  | 48 |  | 17 |  | 48 |  | 16 |  | Confusion | Psychic | Special |
|  | 1-2-053 | Bronzong |  | Steel | Psychic |  | 74 |  | 31 |  | 82 | 60 |  | 76 |  | 53 |  | 76 |  | 24 |  | Iron Head | Steel | Physical |
|  | 1-2-054 | Trapinch |  | Ground |  |  | 40 |  | 25 |  | 57 | 55 |  | 27 |  | 27 |  | 27 |  | 0 |  | Sand Tomb | Ground | Physical |
|  | 1-2-055 | Vibrava |  | Ground | Dragon |  | 58 |  | 33 |  | 76 | 51 |  | 38 |  | 38 |  | 38 |  | 51 |  | Dragon Tail | Dragon | Physical |
|  | 1-2-056 | Flygon | 4-Star | Ground | Dragon |  | 90 |  | 39 |  | 111 | 83 |  | 67 |  | 67 |  | 67 |  | 83 |  | Dragon Claw | Dragon | Physical |
|  | 1-2-057 | Wingull |  | Water | Flying |  | 38 |  | 24 |  | 55 | 19 |  | 19 |  | 34 |  | 19 |  | 50 |  | Gust | Flying | Special |
|  | 1-2-058 | Pelipper |  | Water | Flying |  | 70 |  | 32 |  | 80 | 37 |  | 69 |  | 65 |  | 49 |  | 46 |  | Air Slash | Flying | Special |
|  | 1-2-059 | Jynx |  | Ice | Psychic |  | 68 |  | 33 |  | 79 | 35 |  | 24 |  | 74 |  | 62 |  | 62 |  | Avalanche | Ice | Physical |
|  | 1-2-060 | Abra |  | Psychic |  |  | 42 |  | 25 |  | 47 | 13 |  | 9 |  | 57 |  | 32 |  | 50 |  | Hidden Power | Psychic | Special |
|  | 1-2-061 | Kadabra |  | Psychic |  |  | 62 |  | 31 |  | 65 | 26 |  | 22 |  | 79 |  | 48 |  | 70 |  | Psybeam | Psychic | Special |
|  | 1-2-062 | Alakazam | 4-Star | Psychic |  |  | 88 |  | 37 |  | 87 | 41 |  | 37 |  | 106 |  | 75 |  | 93 |  | Psychic | Psychic | Special |
|  | 1-2-063 | Drifloon |  | Ghost | Flying |  | 58 |  | 31 |  | 96 | 36 |  | 25 |  | 42 |  | 32 |  | 48 |  | Astonish | Ghost | Physical |
|  | 1-2-064 | Drifblim | 4-Star | Ghost | Flying |  | 88 |  | 37 |  | 158 | 64 |  | 37 |  | 71 |  | 44 |  | 64 |  | Hex | Ghost | Special |
|  | 1-2-065 | Mareanie |  | Poison | Water |  | 54 |  | 33 |  | 76 | 39 |  | 45 |  | 33 |  | 39 |  | 34 |  | Venoshock | Poison | Special |
|  | 1-2-066 | Togedemaru |  | Electric | Steel |  | 68 |  | 32 |  | 83 | 67 |  | 45 |  | 30 |  | 51 |  | 66 |  | Spark | Electric | Physical |
|  | 1-2-067 | Ralts |  | Psychic | Fairy |  | 32 |  | 25 |  | 49 | 17 |  | 17 |  | 29 |  | 22 |  | 25 |  | Disarming Voice | Fairy | Special |
|  | 1-2-068 | Kirlia |  | Psychic | Fairy |  | 50 |  | 33 |  | 68 | 27 |  | 27 |  | 47 |  | 41 |  | 38 |  | Draining Kiss | Fairy | Special |
|  | 1-2-069 | Gardevoir | 4-Star | Psychic | Fairy |  | 88 |  | 37 |  | 102 | 55 |  | 55 |  | 102 |  | 94 |  | 67 |  | Psyshock | Psychic | Special |
|  | 1-2-070 | Onix |  | Rock | Ground |  | 64 |  | 33 |  | 66 | 34 |  | 110 |  | 22 |  | 34 |  | 51 |  | Bulldoze | Ground | Physical |
|  | 1-3-001 | Lugia | 6-Star | Psychic | Flying |  | 162 |  | 53 |  | 175 | 100 |  | 142 |  | 100 |  | 168 |  | 121 |  | Aeroblast | Flying | Special |
|  | 1-3-002 | Ho-Oh | 6-Star | Fire | Flying |  | 172 |  | 57 |  | 187 | 153 |  | 107 |  | 130 |  | 180 |  | 107 |  | Sacred Fire | Fire | Physical |
|  | 1-3-003 | Solgaleo | 6-Star | Psychic | Steel |  | 172 |  | 57 |  | 223 | 161 |  | 126 |  | 133 |  | 106 |  | 115 |  | Sunsteel Strike | Steel | Physical |
|  | 1-3-004 | Lunala | 6-Star | Psychic | Ghost |  | 172 |  | 57 |  | 223 | 133 |  | 106 |  | 161 |  | 126 |  | 115 |  | Moongeist Beam | Ghost | Special |
|  | 1-3-005 | Greninja | 6-Star | Water | Dark | Z-Move | 142 |  | 53 |  | 151 | 115 |  | 82 |  | 124 |  | 87 |  | 146 |  | Dark Pulse | Dark | Special |
|  | 1-3-006 | Zeraora | 6-Star | Electric |  | Z-Move | 150 |  | 53 |  | 170 | 134 |  | 92 |  | 123 |  | 97 |  | 170 |  | Plasma Fists | Electric | Physical |
|  | 1-3-007 | Eternatus | 6-Star | Poison | Dragon |  | 168 |  | 54 |  | 215 | 96 |  | 107 |  | 161 |  | 107 |  | 145 |  | Dynamax Cannon | Dragon | Special |
|  | 1-3-008 | Keldeo(Resolute) | 6-Star | Water | Fighting | Dynamax | 158 | 166 / 174 / 182 | 60 |  | 179 | 91 |  | 113 |  | 159 |  | 113 |  | 134 |  | Secret Sword | Fighting | Special |
|  | 1-3-009 | Grimmsnarl | 6-Star | Dark | Fairy | Dynamax | 128 | 135 / 141 / 148 | 54 |  | 166 | 134 |  | 75 |  | 107 |  | 86 |  | 69 |  | Spirit Break | Fairy | Physical |
|  | 1-3-010 | Zygarde(Complete) | 6-Star | Dragon | Ground |  | 164 |  | 52 |  | 286 | 109 |  | 130 |  | 94 |  | 103 |  | 103 |  | Land's Wrath | Ground | Physical |
|  | 1-3-011 | Sceptile(Mega) | 5-Star | Grass |  | Mega Evolution | 128 | 139 | 52 |  | 134 | 93 | 119 | 72 | 83 | 114 | 155 | 93 |  | 129 | 155 | Leaf Storm | Grass | Special |
|  | 1-3-012 | Blaziken(Mega) | 5-Star | Fire | Fighting | Mega Evolution | 128 | 140 | 52 |  | 145 | 129 | 171 | 77 | 88 | 119 | 140 | 77 | 88 | 88 | 109 | Focus Blast | Fighting | Special |
|  | 1-3-013 | Swampert(Mega) | 5-Star | Water | Ground | Mega Evolution | 128 | 139 | 52 |  | 166 | 119 | 161 | 98 | 119 | 93 | 103 | 98 | 119 | 67 | 77 | Earthquake | Ground | Physical |
|  | 1-3-014 | Chandelure | 5-Star | Ghost | Fire | Z-Move | 107 |  | 54 |  | 128 | 64 |  | 102 |  | 161 |  | 102 |  | 91 |  | Shadow Ball | Ghost | Special |
|  | 1-3-015 | Grimmsnarl | 5-Star | Dark | Fairy | Dynamax | 122 | 128 / 135 / 141 | 52 |  | 160 | 129 |  | 72 |  | 103 |  | 83 |  | 67 |  | Spirit Break | Fairy | Physical |
|  | 1-3-016 | Nidoqueen | 5-Star | Poison | Ground | Dynamax | 122 | 128 / 135 / 141 | 52 |  | 155 | 100 |  | 95 |  | 83 |  | 93 |  | 84 |  | Earth Power | Ground | Special |
|  | 1-3-017 | Pidgeot | 5-Star | Normal | Flying | Tag Move | 100 |  | 51 |  | 150 | 89 |  | 84 |  | 79 |  | 79 |  | 112 |  | Hurricane | Flying | Special |
|  | 1-3-018 | Coalossal | 5-Star | Rock | Fire | Tag Move | 110 |  | 54 |  | 182 | 91 |  | 134 |  | 91 |  | 102 |  | 37 |  | Rock Blast | Rock | Physical |
|  | 1-3-019 | Pikachu | 5-Star | Electric |  | Tag Move | 100 |  | 50 |  | 125 | 79 |  | 59 |  | 73 |  | 73 |  | 127 |  | Thunderbolt | Electric | Special |
|  | 1-3-020 | Alakazam | 5-Star | Psychic |  | Tag Move | 120 |  | 52 |  | 119 | 57 |  | 51 |  | 145 |  | 103 |  | 129 |  | Psychic | Psychic | Special |
|  | 1-3-021 | Regirock | 5-Star | Rock |  |  | 140 |  | 53 |  | 147 | 111 |  | 217 |  | 58 |  | 111 |  | 58 |  | Stone Edge | Rock | Physical |
|  | 1-3-022 | Regice | 5-Star | Ice |  |  | 140 |  | 53 |  | 147 | 58 |  | 111 |  | 111 |  | 217 |  | 58 |  | Ice Beam | Ice | Special |
|  | 1-3-023 | Registeel | 5-Star | Steel |  |  | 140 |  | 53 |  | 147 | 84 |  | 164 |  | 84 |  | 164 |  | 58 |  | Flash Cannon | Steel | Special |
|  | 1-3-024 | Great Tusk | 5-Star | Ground | Fighting |  | 130 |  | 51 |  | 184 | 143 |  | 143 |  | 61 |  | 61 |  | 97 |  | Close Combat | Fighting | Physical |
|  | 1-3-025 | Haxorus | 5-Star | Dragon |  |  | 120 |  | 51 |  | 138 | 154 |  | 96 |  | 60 |  | 76 |  | 103 |  | Outrage | Dragon | Physical |
|  | 1-3-026 | Armarouge | 4-Star | Fire | Psychic |  | 112 |  | 48 |  | 134 | 60 |  | 97 |  | 120 |  | 78 |  | 74 |  | Psybeam | Psychic | Special |
|  | 1-3-027 | Arctibax |  | Dragon | Ice |  | 98 |  | 48 |  | 144 | 90 |  | 88 |  | 48 |  | 67 |  | 64 |  | Icy Wind | Ice | Special |
|  | 1-3-028 | Pawmo |  | Electric | Fighting |  | 84 |  | 40 |  | 117 | 78 |  | 44 |  | 54 |  | 44 |  | 88 |  | Arm Thrust | Fighting | Physical |
|  | 1-3-029 | Treecko |  | Grass |  |  | 44 |  | 24 |  | 56 | 28 |  | 23 |  | 38 |  | 33 |  | 41 |  | Leafage | Grass | Physical |
|  | 1-3-030 | Grovyle |  | Grass |  |  | 72 |  | 36 |  | 82 | 51 |  | 37 |  | 66 |  | 51 |  | 73 |  | Assurance | Dark | Physical |
|  | 1-3-031 | Sceptile | 4-Star | Grass |  |  | 104 |  | 42 |  | 110 | 76 |  | 59 |  | 93 |  | 76 |  | 105 |  | Mega Drain | Grass | Special |
|  | 1-3-032 | Torchic |  | Fire |  |  | 44 |  | 24 |  | 59 | 36 |  | 25 |  | 41 |  | 31 |  | 28 |  | Ember | Fire | Special |
|  | 1-3-033 | Combusken |  | Fire | Fighting |  | 72 |  | 36 |  | 89 | 66 |  | 48 |  | 66 |  | 48 |  | 44 |  | Flame Charge | Fire | Physical |
|  | 1-3-034 | Blaziken | 4-Star | Fire | Fighting |  | 104 |  | 42 |  | 119 | 105 |  | 63 |  | 97 |  | 63 |  | 72 |  | Fire Punch | Fire | Physical |
|  | 1-3-035 | Mudkip |  | Water |  |  | 44 |  | 24 |  | 62 | 41 |  | 31 |  | 31 |  | 31 |  | 25 |  | Water Gun | Water | Special |
|  | 1-3-036 | Marshtomp |  | Water | Ground |  | 72 |  | 36 |  | 96 | 66 |  | 55 |  | 48 |  | 55 |  | 41 |  | Mud-Shot | Ground | Special |
|  | 1-3-037 | Swampert | 4-Star | Water | Ground |  | 104 |  | 42 |  | 136 | 97 |  | 80 |  | 70 |  | 80 |  | 55 |  | Bulldoze | Ground | Physical |
|  | 1-3-038 | Magby |  | Fire |  |  | 50 |  | 26 |  | 59 | 44 |  | 24 |  | 41 |  | 33 |  | 48 |  | Ember | Fire | Special |
|  | 1-3-039 | Magmar |  | Fire |  |  | 82 |  | 35 |  | 90 | 71 |  | 44 |  | 75 |  | 64 |  | 70 |  | Flame Wheel | Fire | Physical |
|  | 1-3-040 | Magmortar | 4-Star | Fire |  |  | 104 |  | 41 |  | 112 | 82 |  | 59 |  | 107 |  | 82 |  | 73 |  | Fire Punch | Fire | Physical |
|  | 1-3-041 | Koffing |  | Poison |  |  | 48 |  | 26 |  | 56 | 38 |  | 54 |  | 36 |  | 28 |  | 23 |  | Clear Smog | Poison | Special |
|  | 1-3-042 | Weezing(Galarian) | 4-Star | Poison | Fairy |  | 100 |  | 43 |  | 108 | 82 |  | 108 |  | 78 |  | 65 |  | 56 |  | Sludge | Poison | Special |
|  | 1-3-043 | Beldum |  | Steel | Psychic |  | 42 |  | 26 |  | 56 | 33 |  | 46 |  | 23 |  | 36 |  | 20 |  | Tackle | Normal | Physical |
|  | 1-3-044 | Metang |  | Steel | Psychic |  | 76 |  | 37 |  | 91 | 60 |  | 79 |  | 45 |  | 64 |  | 42 |  | Metal Claw | Steel | Physical |
|  | 1-3-045 | Metagross | 4-Star | Steel | Psychic |  | 118 |  | 43 |  | 121 | 121 |  | 116 |  | 86 |  | 82 |  | 65 |  | Zen Headbutt | Psychic | Physical |
|  | 1-3-046 | Shinx |  | Electric |  |  | 38 |  | 26 |  | 59 | 38 |  | 22 |  | 25 |  | 22 |  | 28 |  | Thunder Shock | Electric | Special |
|  | 1-3-047 | Luxio |  | Electric |  |  | 68 |  | 37 |  | 91 | 67 |  | 41 |  | 49 |  | 41 |  | 49 |  | Spark | Electric | Physical |
|  | 1-3-048 | Luxray | 4-Star | Electric |  |  | 104 |  | 43 |  | 121 | 108 |  | 72 |  | 86 |  | 72 |  | 65 |  | Discharge | Electric | Special |
|  | 1-3-049 | Blipbug |  | Bug |  |  | 30 |  | 26 |  | 49 | 15 |  | 15 |  | 18 |  | 28 |  | 28 |  | Struggle Bug | Bug | Special |
|  | 1-3-050 | Dottler |  | Bug | Psychic |  | 62 |  | 36 |  | 82 | 30 |  | 62 |  | 41 |  | 69 |  | 26 |  | Confusion | Psychic | Special |
|  | 1-3-051 | Orbeetle | 4-Star | Bug | Psychic |  | 100 |  | 42 |  | 102 | 42 |  | 97 |  | 72 |  | 105 |  | 80 |  | Psybeam | Psychic | Special |
|  | 1-3-052 | Silicobra |  | Ground |  |  | 44 |  | 26 |  | 63 | 34 |  | 44 |  | 23 |  | 31 |  | 28 |  | Bulldoze | Ground | Physical |
|  | 1-3-053 | Sandaconda | 4-Star | Ground |  |  | 106 |  | 44 |  | 117 | 99 |  | 115 |  | 62 |  | 66 |  | 67 |  | Dig | Ground | Physical |
|  | 1-3-054 | Rolycoly |  | Rock |  |  | 36 |  | 26 |  | 51 | 25 |  | 31 |  | 25 |  | 31 |  | 20 |  | Tackle | Normal | Physical |
|  | 1-3-055 | Carkol |  | Rock | Fire |  | 72 |  | 36 |  | 103 | 48 |  | 69 |  | 48 |  | 55 |  | 41 |  | Smack Down | Rock | Physical |
|  | 1-3-056 | Coalossal | 4-Star | Rock | Fire |  | 100 |  | 42 |  | 144 | 72 |  | 105 |  | 72 |  | 80 |  | 30 |  | Ancient Power | Rock | Special |
|  | 1-3-057 | Mudbray |  | Ground |  |  | 72 |  | 38 |  | 101 | 81 |  | 58 |  | 39 |  | 46 |  | 39 |  | Bulldoze | Ground | Physical |
|  | 1-3-058 | Mudsdale | 4-Star | Ground |  |  | 104 |  | 44 |  | 142 | 115 |  | 93 |  | 53 |  | 79 |  | 35 |  | Stomping Tantrum | Ground | Physical |
|  | 1-3-059 | Impidimp |  | Dark | Fairy |  | 40 |  | 26 |  | 59 | 28 |  | 20 |  | 33 |  | 25 |  | 31 |  | Fake Out | Normal | Physical |
|  | 1-3-060 | Morgrem |  | Dark | Fairy |  | 69 |  | 36 |  | 92 | 48 |  | 37 |  | 59 |  | 44 |  | 55 |  | Assurance | Dark | Physical |
|  | 1-3-061 | Grimmsnarl | 4-Star | Dark | Fairy |  | 100 |  | 42 |  | 131 | 105 |  | 59 |  | 84 |  | 68 |  | 55 |  | Sucker Punch | Dark | Physical |
|  | 1-3-062 | Turtonator | 4-Star | Fire | Dragon |  | 98 |  | 43 |  | 104 | 72 |  | 121 |  | 83 |  | 78 |  | 35 |  | Incinerate | Fire | Special |
|  | 1-3-063 | Nickit |  | Dark |  |  | 38 |  | 26 |  | 56 | 19 |  | 19 |  | 29 |  | 32 |  | 31 |  | Quick Attack | Normal | Physical |
|  | 1-3-064 | Thievul | 4-Star | Dark |  |  | 84 |  | 43 |  | 113 | 54 |  | 54 |  | 79 |  | 84 |  | 82 |  | Snarl | Dark | Special |
|  | 1-3-065 | Duskull |  | Ghost |  |  | 42 |  | 26 |  | 46 | 25 |  | 51 |  | 20 |  | 51 |  | 18 |  | Astonish | Ghost | Physical |
|  | 1-3-066 | Dusclops |  | Ghost |  |  | 80 |  | 36 |  | 74 | 55 |  | 98 |  | 48 |  | 98 |  | 23 |  | Shadow Sneak | Ghost | Physical |
|  | 1-3-067 | Dusknoir | 4-Star | Ghost |  |  | 104 |  | 43 |  | 89 | 89 |  | 118 |  | 59 |  | 118 |  | 42 |  | Hex | Ghost | Special |
|  | 1-3-068 | Axew |  | Dragon |  |  | 44 |  | 26 |  | 59 | 50 |  | 36 |  | 20 |  | 25 |  | 34 |  | Scratch | Normal | Physical |
|  | 1-3-069 | Fraxure |  | Dragon |  |  | 72 |  | 35 |  | 91 | 86 |  | 54 |  | 33 |  | 40 |  | 51 |  | Breaking Swipe | Dragon | Physical |
|  | 1-3-070 | Haxorus | 4-Star | Dragon |  |  | 104 |  | 41 |  | 113 | 125 |  | 78 |  | 54 |  | 62 |  | 84 |  | Dual Chop | Dragon | Physical |
|  | 1-4-001 | Dialga | 6-Star | Steel | Dragon |  | 180 |  | 59 |  | 187 | 146 |  | 146 |  | 182 |  | 123 |  | 111 |  | Roar of Time | Dragon | Special |
|  | 1-4-002 | Palkia | 6-Star | Water | Dragon |  | 180 |  | 59 |  | 175 | 146 |  | 123 |  | 182 |  | 146 |  | 123 |  | Spacial Rend | Dragon | Special |
|  | 1-4-003 | Xerneas | 6-Star | Fairy |  | Dynamax | 178 | 187 / 195 / 203 | 59 |  | 217 | 159 |  | 117 |  | 159 |  | 120 |  | 121 |  | Moonblast | Fairy | Special |
|  | 1-4-004 | Yveltal | 6-Star | Dark | Flying | Dynamax | 178 | 187 / 195 / 203 | 59 |  | 217 | 159 |  | 117 |  | 159 |  | 120 |  | 121 |  | Dark Pulse | Dark | Special |
|  | 1-4-005 | Calyrex(Ice Rider) | 6-Star | Psychic | Ice | Dynamax | 182 | 192 / 201 / 210 | 60 |  | 190 | 203 |  | 185 |  | 107 |  | 161 |  | 65 |  | Glacial Lance | Ice | Physical |
|  | 1-4-006 | Calyrex(Shadow Rider) | 6-Star | Psychic | Ghost | Dynamax | 182 | 192 / 201 / 210 | 60 |  | 190 | 107 |  | 101 |  | 203 |  | 125 |  | 185 |  | Astral Barrage | Ghost | Special |
|  | 1-4-007 | Mimikyu | 6-Star | Ghost | Fairy | Z-Move Signature | 142 |  | 59 | Disguise | 133 | 111 |  | 99 |  | 64 |  | 128 |  | 118 |  | Play Rough | Fairy | Physical |
|  | 1-4-008 | Kyurem(White) | 6-Star | Dragon | Ice |  | 180 |  | 59 |  | 213 | 144 |  | 109 |  | 202 |  | 121 |  | 115 |  | Fusion Flare | Fire | Special |
|  | 1-4-009 | Kyurem(Black) | 6-Star | Dragon | Ice |  | 180 |  | 59 |  | 213 | 202 |  | 121 |  | 144 |  | 109 |  | 115 |  | Fusion Bolt | Electric | Physical |
|  | 1-4-010 | Rayquaza | 6-Star | Dragon | Flying |  | 182 |  | 60 |  | 196 | 185 |  | 113 |  | 185 |  | 113 |  | 119 |  | Draco Meteor | Dragon | Special |
|  | 1-4-011 | Aggron(Mega) | 5-Star | Steel | Rock | Mega Evolution | 134 | 149 | 55 |  | 142 | 126 | 159 | 203 | 258 | 71 |  | 71 | 93 | 60 |  | Iron Tail | Steel | Physical |
|  | 1-4-012 | Meganium | 5-Star | Grass |  | Z-Move | 124 |  | 51 |  | 142 | 88 |  | 107 |  | 89 |  | 107 |  | 80 |  | Solar Beam | Grass | Special |
|  | 1-4-013 | Typhlosion | 5-Star | Fire |  | Z-Move | 126 |  | 51 |  | 140 | 90 |  | 84 |  | 116 |  | 91 |  | 107 |  | Flamethrower | Fire | Special |
|  | 1-4-014 | Feraligatr | 5-Star | Water |  | Z-Move | 126 |  | 51 |  | 147 | 112 |  | 107 |  | 85 |  | 89 |  | 84 |  | Hydro Pump | Water | Special |
|  | 1-4-015 | Goodra | 5-Star | Dragon |  | Dynamax | 144 | 150 / 157 / 163 / 171 | 54 |  | 161 | 113 |  | 80 |  | 123 |  | 167 |  | 91 |  | Dragon Pulse | Dragon | Special |
|  | 1-4-016 | Corviknight | 5-Star | Flying | Steel | Dynamax | 122 | 128 / 135 / 141 / 148 | 51 |  | 172 | 100 |  | 120 |  | 63 |  | 98 |  | 78 |  | Brave Bird | Flying | Physical |
|  | 1-4-017 | Runerigus | 5-Star | Ground | Ghost | Tag Move | 120 |  | 53 |  | 124 | 105 |  | 158 |  | 58 |  | 116 |  | 36 |  | Shadow Claw | Ghost | Physical |
|  | 1-4-018 | Sirfetch'd | 5-Star | Fighting |  | Tag Move | 122 |  | 52 |  | 155 | 140 |  | 88 |  | 72 |  | 93 |  | 62 |  | Close Combat | Fighting | Physical |
|  | 1-4-019 | Gardevoir | 5-Star | Psychic | Fairy | Tag Move | 132 |  | 55 |  | 139 | 76 |  | 76 |  | 142 |  | 131 |  | 93 |  | Dazzling Gleam | Fairy | Special |
|  | 1-4-020 | Obstagoon | 5-Star | Dark | Normal | Tag Move | 120 |  | 52 |  | 158 | 98 |  | 110 |  | 67 |  | 89 |  | 103 |  | Throat Chop | Dark | Physical |
|  | 1-4-021 | Regieleki | 5-Star | Electric |  |  | 144 |  | 54 |  | 150 | 113 |  | 59 |  | 113 |  | 59 |  | 221 |  | Thunder Cage | Electric | Special |
|  | 1-4-022 | Regidrago | 5-Star | Dragon |  |  | 144 |  | 54 |  | 280 | 113 |  | 59 |  | 113 |  | 59 |  | 91 |  | Dragon Energy | Dragon | Special |
|  | 1-4-023 | Iron Treads | 5-Star | Ground | Steel |  | 140 |  | 54 |  | 161 | 125 |  | 134 |  | 82 |  | 80 |  | 119 |  | Earthquake | Ground | Physical |
|  | 1-4-024 | Copperajah | 5-Star | Steel |  |  | 118 |  | 51 |  | 185 | 137 |  | 75 |  | 86 |  | 75 |  | 35 |  | Iron Head | Steel | Physical |
|  | 1-4-025 | Talonflame | 5-Star | Fire | Flying |  | 120 |  | 52 |  | 143 | 89 |  | 78 |  | 81 |  | 76 |  | 136 |  | Flare Blitz | Fire | Physical |
|  | 1-4-026 | Ceruledge | 4-Star | Fire | Ghost |  | 114 |  | 47 |  | 127 | 122 |  | 80 |  | 61 |  | 99 |  | 84 |  | Hex | Ghost | Special |
|  | 1-4-027 | Baxcalibur | 4-Star | Dragon | Ice |  | 128 |  | 47 |  | 165 | 141 |  | 91 |  | 75 |  | 85 |  | 86 |  | Ice Fang | Ice | Physical |
|  | 1-4-028 | Chikorita |  | Grass |  |  | 44 |  | 24 |  | 59 | 30 |  | 38 |  | 30 |  | 38 |  | 28 |  | Razor Leaf | Grass | Physical |
|  | 1-4-029 | Bayleef |  | Grass |  |  | 72 |  | 35 |  | 87 | 48 |  | 61 |  | 49 |  | 61 |  | 47 |  | Magical Leaf | Grass | Special |
|  | 1-4-030 | Meganium | 4-Star | Grass |  |  | 102 |  | 41 |  | 116 | 72 |  | 87 |  | 73 |  | 87 |  | 70 |  | Grass Pledge | Grass | Special |
|  | 1-4-031 | Cyndaquil |  | Fire |  |  | 44 |  | 24 |  | 56 | 32 |  | 27 |  | 36 |  | 31 |  | 38 |  | Flame Charge | Fire | Physical |
|  | 1-4-032 | Quilava |  | Fire |  |  | 72 |  | 35 |  | 85 | 49 |  | 45 |  | 61 |  | 50 |  | 61 |  | Flame Wheel | Fire | Physical |
|  | 1-4-033 | Typhlosion | 4-Star | Fire |  |  | 102 |  | 41 |  | 114 | 73 |  | 68 |  | 94 |  | 74 |  | 87 |  | Lava Plume | Fire | Special |
|  | 1-4-034 | Totodile |  | Water |  |  | 44 |  | 24 |  | 62 | 38 |  | 38 |  | 27 |  | 29 |  | 27 |  | Water Gun | Water | Special |
|  | 1-4-035 | Croconaw |  | Water |  |  | 72 |  | 35 |  | 90 | 61 |  | 61 |  | 46 |  | 49 |  | 45 |  | Water Pulse | Water | Special |
|  | 1-4-036 | Feraligatr | 4-Star | Water |  |  | 102 |  | 41 |  | 120 | 91 |  | 87 |  | 69 |  | 73 |  | 68 |  | Waterfall | Water | Physical |
|  | 1-4-037 | Dreepy |  | Dragon | Ghost |  | 40 |  | 26 |  | 50 | 36 |  | 20 |  | 25 |  | 20 |  | 47 |  | Astonish | Ghost | Physical |
|  | 1-4-038 | Drakloak |  | Dragon | Ghost |  | 74 |  | 37 |  | 97 | 64 |  | 42 |  | 49 |  | 42 |  | 80 |  | Breaking Swipe | Dragon | Physical |
|  | 1-4-039 | Dragapult | 4-Star | Dragon | Ghost |  | 110 |  | 43 |  | 128 | 108 |  | 69 |  | 91 |  | 69 |  | 127 |  | Dragon Claw | Dragon | Physical |
|  | 1-4-040 | Sinistea |  | Ghost |  |  | 44 |  | 26 |  | 50 | 28 |  | 28 |  | 43 |  | 33 |  | 31 |  | Astonish | Ghost | Physical |
|  | 1-4-041 | Polteageist | 4-Star | Ghost |  |  | 92 |  | 41 |  | 100 | 58 |  | 58 |  | 114 |  | 98 |  | 62 |  | Hex | Ghost | Special |
|  | 1-4-042 | Drampa |  | Normal | Dragon |  | 82 |  | 35 |  | 99 | 47 |  | 64 |  | 99 |  | 68 |  | 30 |  | Dragon Breath | Dragon | Special |
|  | 1-4-043 | Elekid |  | Electric |  |  | 50 |  | 26 |  | 59 | 37 |  | 24 |  | 38 |  | 33 |  | 54 |  | Thunder Shock | Electric | Special |
|  | 1-4-044 | Electabuzz |  | Electric |  |  | 84 |  | 37 |  | 95 | 66 |  | 47 |  | 75 |  | 67 |  | 82 |  | Shock Wave | Electric | Special |
|  | 1-4-045 | Electivire | 4-Star | Electric |  |  | 108 |  | 43 |  | 117 | 110 |  | 62 |  | 86 |  | 78 |  | 86 |  | Thunder Punch | Electric | Physical |
|  | 1-4-046 | Vulpix(Alolan) |  | Ice |  |  | 56 |  | 26 |  | 73 | 34 |  | 33 |  | 41 |  | 51 |  | 51 |  | Icy Wind | Ice | Special |
|  | 1-4-047 | Ninetales(Alolan) | 4-Star | Ice | Fairy |  | 100 |  | 42 |  | 113 | 61 |  | 68 |  | 73 |  | 89 |  | 96 |  | Aurora Beam | Ice | Special |
|  | 1-4-048 | Wooloo |  | Normal |  |  | 40 |  | 26 |  | 57 | 25 |  | 33 |  | 25 |  | 28 |  | 29 |  | Tackle | Normal | Physical |
|  | 1-4-049 | Dubwool | 4-Star | Normal |  |  | 88 |  | 41 |  | 110 | 70 |  | 87 |  | 54 |  | 78 |  | 77 |  | Headbutt | Normal | Physical |
|  | 1-4-050 | Deino |  | Dark | Dragon |  | 44 |  | 26 |  | 63 | 38 |  | 31 |  | 28 |  | 31 |  | 24 |  | Tackle | Normal | Physical |
|  | 1-4-051 | Zweilous |  | Dark | Dragon |  | 74 |  | 37 |  | 100 | 67 |  | 56 |  | 53 |  | 56 |  | 47 |  | Assurance | Dark | Physical |
|  | 1-4-052 | Hydreigon | 4-Star | Dark | Dragon |  | 114 |  | 44 |  | 132 | 95 |  | 82 |  | 112 |  | 82 |  | 89 |  | Dual Chop | Dragon | Physical |
|  | 1-4-053 | Woobat |  | Psychic | Flying |  | 46 |  | 26 |  | 69 | 28 |  | 27 |  | 33 |  | 27 |  | 42 |  | Confusion | Psychic | Special |
|  | 1-4-054 | Swoobat |  | Psychic | Flying |  | 76 |  | 36 |  | 94 | 46 |  | 44 |  | 60 |  | 44 |  | 87 |  | Air Cutter | Flying | Special |
|  | 1-4-055 | Machop |  | Fighting |  |  | 44 |  | 26 |  | 72 | 46 |  | 31 |  | 23 |  | 23 |  | 23 |  | Low Sweep | Fighting | Physical |
|  | 1-4-056 | Machoke |  | Fighting |  |  | 72 |  | 36 |  | 103 | 77 |  | 55 |  | 41 |  | 48 |  | 37 |  | Vital Throw | Fighting | Physical |
|  | 1-4-057 | Machamp | 4-Star | Fighting |  |  | 100 |  | 42 |  | 127 | 114 |  | 72 |  | 59 |  | 76 |  | 51 |  | Strength | Normal | Physical |
|  | 1-4-058 | Murkrow |  | Dark | Flying |  | 72 |  | 36 |  | 89 | 66 |  | 35 |  | 66 |  | 35 |  | 70 |  | Pursuit | Dark | Physical |
|  | 1-4-059 | Honchkrow | 4-Star | Dark | Flying |  | 100 |  | 42 |  | 136 | 110 |  | 48 |  | 93 |  | 48 |  | 64 |  | Night Slash | Dark | Physical |
|  | 1-4-060 | Hatenna |  | Psychic |  |  | 40 |  | 26 |  | 57 | 20 |  | 28 |  | 34 |  | 32 |  | 25 |  | Disarming Voice | Fairy | Special |
|  | 1-4-061 | Hattrem |  | Psychic |  |  | 70 |  | 36 |  | 91 | 35 |  | 54 |  | 70 |  | 60 |  | 42 |  | Confusion | Psychic | Special |
|  | 1-4-062 | Hatterene | 4-Star | Psychic | Fairy |  | 104 |  | 44 |  | 104 | 84 |  | 88 |  | 124 |  | 95 |  | 30 |  | Psyshock | Psychic | Special |
|  | 1-4-063 | Swinub |  | Ice | Ground |  | 38 |  | 24 |  | 62 | 31 |  | 25 |  | 20 |  | 20 |  | 31 |  | Powder Snow | Ice | Special |
|  | 1-4-064 | Piloswine |  | Ice | Ground |  | 82 |  | 36 |  | 124 | 81 |  | 65 |  | 50 |  | 50 |  | 43 |  | Icy Wind | Ice | Special |
|  | 1-4-065 | Mamoswine | 4-Star | Ice | Ground |  | 108 |  | 44 |  | 150 | 119 |  | 75 |  | 66 |  | 57 |  | 75 |  | Ice Fang | Ice | Physical |
|  | 1-4-066 | Fletchling |  | Normal | Flying |  | 40 |  | 26 |  | 59 | 31 |  | 27 |  | 25 |  | 24 |  | 37 |  | Peck | Flying | Physical |
|  | 1-4-067 | Fletchinder |  | Fire | Flying |  | 68 |  | 36 |  | 90 | 57 |  | 44 |  | 45 |  | 42 |  | 65 |  | Acrobatics | Flying | Physical |
|  | 1-4-068 | Talonflame | 4-Star | Fire | Flying |  | 100 |  | 43 |  | 117 | 73 |  | 64 |  | 67 |  | 62 |  | 110 |  | Aerial Ace | Flying | Physical |
|  | 1-4-069 | Noibat |  | Flying | Dragon |  | 44 |  | 26 |  | 56 | 20 |  | 23 |  | 28 |  | 25 |  | 33 |  | Gust | Flying | Special |
|  | 1-4-070 | Noivern | 4-Star | Flying | Dragon |  | 100 |  | 41 |  | 120 | 62 |  | 70 |  | 84 |  | 70 |  | 105 |  | Dragon Claw | Dragon | Physical |
|  | 2-1-001 | Rillaboom(Gigantamax) | 6-Star | Grass |  | Gigantamax | 158 | 166 / 174 / 182 | 64 |  | 202 | 145 |  | 120 |  | 81 |  | 94 |  | 113 |  | Drum Beating | Grass | Physical |
|  | 2-1-002 | Cinderace(Gigantamax) | 6-Star | Fire |  | Gigantamax | 158 | 166 / 174 / 182 | 64 |  | 176 | 153 |  | 101 |  | 88 |  | 101 |  | 157 |  | Pyro Ball | Fire | Physical |
|  | 2-1-003 | Inteleon(Gigantamax) | 6-Star | Water |  | Gigantamax | 158 | 166 / 174 / 182 | 64 |  | 163 | 113 |  | 88 |  | 165 |  | 88 |  | 158 |  | Snipe Shot | Water | Special |
|  | 2-1-004 | Pikachu(Gigantamax) | 6-Star | Electric |  | Gigantamax | 122 | 129 / 135 / 141 | 70 |  | 142 | 90 |  | 67 |  | 83 |  | 83 |  | 145 |  | Thunderbolt | Electric | Special |
|  | 2-1-005 | Raikou | 6-Star | Electric |  | Double Move | 170 |  | 60 |  | 192 | 115 |  | 102 |  | 154 |  | 135 |  | 154 |  | Thunder | Electric | Special |
|  | 2-1-006 | Entei | 6-Star | Fire |  | Double Move | 170 |  | 60 |  | 224 | 154 |  | 115 |  | 122 |  | 102 |  | 135 |  | Fire Blast | Fire | Special |
|  | 2-1-007 | Suicune | 6-Star | Water |  | Double Move | 170 |  | 60 |  | 205 | 102 |  | 154 |  | 122 |  | 154 |  | 115 |  | Hydro Pump | Water | Special |
|  | 2-1-008 | Mewtwo | 6-Star | Psychic |  | Dynamax | 194 | 194 / 203 / 212 | 61 |  | 200 | 139 |  | 114 |  | 192 |  | 114 |  | 163 |  | Psystrike | Psychic | Special |
|  | 2-1-009 | Arceus | 6-Star | Normal |  |  | 204 |  | 60 |  | 214 | 149 |  | 149 |  | 149 |  | 149 |  | 149 |  | Judgment | Normal | Special |
|  | 2-1-010 | Zoroark(Hisuian) | 6-Star | Normal | Ghost |  | 150 |  | 54 |  | 144 | 133 |  | 81 |  | 165 |  | 81 |  | 145 |  | Bitter Malice | Ghost | Special |
|  | 2-1-011 | Pinsir(Mega) |  | Bug |  | Mega Evolution | 126 | 138 | 54 |  | 134 | 140 | 172 | 113 | 134 | 64 | 75 | 80 | 102 | 95 | 118 | Superpower | Fighting | Physical |
|  | 2-1-012 | Heracross(Mega) |  | Bug | Fighting | Mega Evolution | 126 | 138 | 54 |  | 150 | 140 | 204 | 86 | 129 | 48 |  | 107 | 118 | 95 | 86 | Megahorn | Bug | Physical |
|  | 2-1-013 | Keldeo |  | Water | Fighting | Z-Move | 142 |  | 54 |  | 162 | 82 |  | 102 |  | 144 |  | 102 |  | 121 |  | Sacred Sword | Fighting | Physical |
|  | 2-1-014 | Weavile |  | Dark | Ice | Dynamax | 128 | 135 / 141 / 148 | 54 |  | 139 | 134 |  | 75 |  | 53 |  | 96 |  | 140 |  | Throat Chop | Dark | Physical |
|  | 2-1-015 | Sirfetch'd |  | Fighting |  | Dynamax | 136 | 143 / 150 / 157 | 50 |  | 139 | 161 |  | 115 |  | 83 |  | 100 |  | 80 |  | Meteor Assault | Fighting | Physical |
|  | 2-1-016 | Nidoking |  | Poison | Ground | Dynamax | 128 | 135 / 141 / 148 | 50 |  | 154 | 117 |  | 89 |  | 98 |  | 87 |  | 98 |  | Earth Power | Ground | Special |
|  | 2-1-017 | Flygon |  | Ground | Dragon | Tag Move | 144 |  | 60 |  | 166 | 125 |  | 101 |  | 101 |  | 101 |  | 125 |  | Earthquake | Ground | Physical |
|  | 2-1-018 | Arctovish |  | Water | Ice | Tag Move | 130 |  | 56 |  | 172 | 109 |  | 121 |  | 97 |  | 109 |  | 68 |  | Hydro Pump | Water | Special |
|  | 2-1-019 | Electivire |  | Electric |  | Chain Attack | 144 |  | 56 |  | 155 | 147 |  | 82 |  | 115 |  | 103 |  | 115 |  | Wild Charge | Electric | Physical |
|  | 2-1-020 | Magmortar |  | Fire |  | Chain Attack | 144 |  | 56 |  | 155 | 115 |  | 82 |  | 150 |  | 115 |  | 101 |  | Fire Blast | Fire | Special |
|  | 2-1-021 | Toxtricity(Amped) |  | Electric | Poison | Chain Attack | 134 |  | 56 |  | 155 | 118 |  | 86 |  | 137 |  | 86 |  | 92 |  | Overdrive | Electric | Special |
|  | 2-1-022 | Toxtricity(Low Key) |  | Electric | Poison | Chain Attack | 134 |  | 56 |  | 155 | 118 |  | 86 |  | 137 |  | 86 |  | 92 |  | Sludge Wave | Poison | Special |
|  | 2-1-023 | Great Tusk |  | Ground | Fighting |  | 146 |  | 56 |  | 194 | 151 |  | 151 |  | 64 |  | 64 |  | 102 |  | Close Combat | Fighting | Physical |
|  | 2-1-024 | Vikavolt |  | Bug | Electric |  | 120 |  | 55 |  | 149 | 82 |  | 104 |  | 164 |  | 87 |  | 52 |  | X-Scissor | Bug | Physical |
|  | 2-1-025 | Abomasnow |  | Grass | Ice |  | 126 |  | 55 |  | 164 | 106 |  | 87 |  | 106 |  | 98 |  | 71 |  | Blizzard | Ice | Special |
|  | 2-1-026 | Sprigatito |  | Grass |  |  | 48 |  | 28 |  | 60 | 39 |  | 35 |  | 30 |  | 30 |  | 41 |  | Leafage | Grass | Physical |
|  | 2-1-027 | Fuecoco |  | Fire |  |  | 48 |  | 28 |  | 75 | 30 |  | 38 |  | 40 |  | 27 |  | 25 |  | Ember | Fire | Special |
|  | 2-1-028 | Quaxly |  | Water |  |  | 48 |  | 28 |  | 68 | 41 |  | 30 |  | 33 |  | 30 |  | 33 |  | Water Gun | Water | Special |
|  | 2-1-029 | Charcadet |  | Fire |  |  | 40 |  | 28 |  | 60 | 33 |  | 27 |  | 33 |  | 27 |  | 24 |  | Ember | Fire | Special |
|  | 2-1-030 | Frigibax |  | Dragon | Ice |  | 48 |  | 28 |  | 74 | 47 |  | 30 |  | 24 |  | 30 |  | 35 |  | Tackle | Normal | Physical |
|  | 2-1-031 | Pawmi |  | Electric |  |  | 40 |  | 28 |  | 63 | 33 |  | 16 |  | 27 |  | 19 |  | 38 |  | Scratch | Normal | Physical |
|  | 2-1-032 | Grookey |  | Grass |  |  | 46 |  | 27 |  | 64 | 40 |  | 32 |  | 26 |  | 26 |  | 40 |  | Branch Poke | Grass | Physical |
|  | 2-1-033 | Thwackey |  | Grass |  |  | 84 |  | 41 |  | 108 | 74 |  | 62 |  | 50 |  | 54 |  | 70 |  | Razor Leaf | Grass | Physical |
|  | 2-1-034 | Rillaboom | 4-Star | Grass |  |  | 116 |  | 47 |  | 151 | 122 |  | 89 |  | 61 |  | 70 |  | 84 |  | Bullet Seed | Grass | Physical |
|  | 2-1-035 | Scorbunny |  | Fire |  |  | 46 |  | 27 |  | 64 | 43 |  | 26 |  | 26 |  | 26 |  | 42 |  | Ember | Fire | Special |
|  | 2-1-036 | Raboot |  | Fire |  |  | 84 |  | 41 |  | 104 | 75 |  | 54 |  | 50 |  | 54 |  | 82 |  | Flame Charge | Fire | Physical |
|  | 2-1-037 | Cinderace | 4-Star | Fire |  |  | 116 |  | 47 |  | 132 | 114 |  | 75 |  | 66 |  | 75 |  | 116 |  | Blaze Kick | Fire | Physical |
|  | 2-1-038 | Sobble |  | Water |  |  | 46 |  | 27 |  | 64 | 26 |  | 26 |  | 42 |  | 26 |  | 42 |  | Water Gun | Water | Special |
|  | 2-1-039 | Drizzile |  | Water |  |  | 84 |  | 41 |  | 104 | 54 |  | 50 |  | 82 |  | 50 |  | 78 |  | Water Pulse | Water | Special |
|  | 2-1-040 | Inteleon | 4-Star | Water |  |  | 116 |  | 47 |  | 122 | 84 |  | 66 |  | 122 |  | 66 |  | 117 |  | Dive | Water | Physical |
|  | 2-1-041 | Snover |  | Grass | Ice |  | 66 |  | 39 |  | 95 | 53 |  | 44 |  | 53 |  | 51 |  | 36 |  | Ice Shard | Ice | Physical |
|  | 2-1-042 | Abomasnow | 4-Star | Grass | Ice |  | 104 |  | 45 |  | 136 | 87 |  | 72 |  | 87 |  | 81 |  | 59 |  | Ice Punch | Ice | Physical |
|  | 2-1-043 | Drifloon |  | Ghost | Flying |  | 68 |  | 38 |  | 116 | 43 |  | 30 |  | 50 |  | 38 |  | 58 |  | Astonish | Ghost | Physical |
|  | 2-1-044 | Drifblim | 4-Star | Ghost | Flying |  | 104 |  | 44 |  | 186 | 75 |  | 43 |  | 84 |  | 52 |  | 75 |  | Hex | Ghost | Special |
|  | 2-1-045 | Zorua(Hisuian) |  | Normal | Ghost |  | 68 |  | 41 |  | 79 | 54 |  | 37 |  | 74 |  | 37 |  | 62 |  | Shadow Sneak | Ghost | Physical |
|  | 2-1-046 | Zoroark(Hisuian) | 4-Star | Normal | Ghost |  | 112 |  | 47 |  | 108 | 99 |  | 61 |  | 122 |  | 61 |  | 108 |  | Swift | Normal | Special |
|  | 2-1-047 | Lileep |  | Rock | Grass |  | 74 |  | 42 |  | 107 | 39 |  | 69 |  | 56 |  | 78 |  | 24 |  | Mega Drain | Grass | Special |
|  | 2-1-048 | Cradily | 4-Star | Rock | Grass |  | 112 |  | 48 |  | 140 | 82 |  | 98 |  | 82 |  | 107 |  | 46 |  | Giga Drain | Grass | Special |
|  | 2-1-049 | Farfetch'd(Galarian) |  | Fighting |  |  | 52 |  | 27 |  | 65 | 56 |  | 34 |  | 36 |  | 38 |  | 34 |  | Rock Smash | Fighting | Physical |
|  | 2-1-050 | Sirfetch'd | 4-Star | Fighting |  |  | 112 |  | 47 |  | 115 | 131 |  | 94 |  | 68 |  | 82 |  | 66 |  | Brick Break | Fighting | Physical |
|  | 2-1-051 | Drilbur |  | Ground |  |  | 68 |  | 41 |  | 100 | 74 |  | 37 |  | 29 |  | 41 |  | 60 |  | Metal Claw | Steel | Physical |
|  | 2-1-052 | Excadrill | 4-Star | Ground | Steel |  | 112 |  | 47 |  | 160 | 131 |  | 61 |  | 52 |  | 66 |  | 87 |  | Drill Run | Ground | Physical |
|  | 2-1-053 | Magby |  | Fire |  |  | 50 |  | 27 |  | 61 | 45 |  | 24 |  | 42 |  | 34 |  | 49 |  | Ember | Fire | Special |
|  | 2-1-054 | Magmar |  | Fire |  |  | 82 |  | 41 |  | 104 | 82 |  | 51 |  | 87 |  | 74 |  | 81 |  | Flame Wheel | Fire | Physical |
|  | 2-1-055 | Magmortar | 4-Star | Fire |  |  | 110 |  | 47 |  | 127 | 94 |  | 67 |  | 122 |  | 94 |  | 83 |  | Fire Punch | Fire | Physical |
|  | 2-1-056 | Growlithe |  | Fire |  |  | 74 |  | 42 |  | 98 | 63 |  | 42 |  | 63 |  | 47 |  | 55 |  | Ember | Fire | Special |
|  | 2-1-057 | Arcanine | 4-Star | Fire |  |  | 122 |  | 48 |  | 144 | 110 |  | 81 |  | 101 |  | 81 |  | 96 |  | Fire Fang | Fire | Physical |
|  | 2-1-058 | Toxel |  | Electric | Poison |  | 38 |  | 27 |  | 58 | 25 |  | 23 |  | 34 |  | 23 |  | 26 |  | Nuzzle | Electric | Physical |
|  | 2-1-059 | Toxtricity(Amped) | 4-Star | Electric | Poison |  | 110 |  | 47 |  | 127 | 97 |  | 70 |  | 112 |  | 70 |  | 75 |  | Discharge | Electric | Special |
|  | 2-1-060 | Toxtricity(Low Key) | 4-Star | Electric | Poison |  | 110 |  | 47 |  | 127 | 97 |  | 70 |  | 112 |  | 70 |  | 75 |  | Poison Jab | Poison | Physical |
|  | 2-1-061 | Gossifleur |  | Grass |  |  | 54 |  | 39 |  | 80 | 36 |  | 51 |  | 36 |  | 51 |  | 12 |  | Leafage | Grass | Physical |
|  | 2-1-062 | Eldegoss | 4-Star | Grass |  |  | 98 |  | 45 |  | 109 | 50 |  | 86 |  | 77 |  | 113 |  | 59 |  | Leaf Tornado | Grass | Special |
|  | 2-1-063 | Sneasel |  | Dark | Ice |  | 60 |  | 38 |  | 89 | 77 |  | 46 |  | 31 |  | 62 |  | 92 |  | Payback | Dark | Physical |
|  | 2-1-064 | Weavile | 4-Star | Dark | Ice |  | 108 |  | 44 |  | 115 | 110 |  | 62 |  | 44 |  | 79 |  | 115 |  | Night Slash | Dark | Physical |
|  | 2-1-065 | Nidoran Female |  | Poison |  |  | 42 |  | 27 |  | 66 | 30 |  | 33 |  | 26 |  | 26 |  | 27 |  | Poison Sting | Poison | Physical |
|  | 2-1-066 | Nidorina |  | Poison |  |  | 70 |  | 41 |  | 110 | 57 |  | 61 |  | 51 |  | 51 |  | 52 |  | Bite | Dark | Physical |
|  | 2-1-067 | Nidoqueen | 4-Star | Poison | Ground |  | 114 |  | 48 |  | 144 | 93 |  | 88 |  | 77 |  | 86 |  | 77 |  | Crunch | Dark | Physical |
|  | 2-1-068 | Nidoran Male |  | Poison |  |  | 42 |  | 27 |  | 61 | 35 |  | 26 |  | 26 |  | 26 |  | 32 |  | Poison Sting | Poison | Physical |
|  | 2-1-069 | Nidorino |  | Poison |  |  | 70 |  | 41 |  | 103 | 65 |  | 52 |  | 51 |  | 51 |  | 59 |  | Double Kick | Fighting | Physical |
|  | 2-1-070 | Nidoking | 4-Star | Poison | Ground |  | 114 |  | 48 |  | 135 | 102 |  | 78 |  | 86 |  | 77 |  | 86 |  | Poison Jab | Poison | Physical |
|  | 2-2-001 | Zacian (Crowned) | 6-Star | Fairy | Steel | Double Move | 206 |  | 60 |  | 194 | 228 |  | 154 |  | 109 |  | 154 |  | 197 |  | Behemoth Blade | Steel | Physical |
|  | 2-2-002 | Zamazenta (Crowned) | 6-Star | Fighting | Steel | Double Move | 206 |  | 61 |  | 194 | 174 |  | 193 |  | 109 |  | 193 |  | 171 |  | Behemoth Bash | Steel | Physical |
|  | 2-2-003 | Duraludon (Gigantamax) | 6-Star | Steel | Dragon | Gigantamax | 158 | 166 / 174 / 182 / 189 | 60 |  | 166 | 128 |  | 154 |  | 161 |  | 70 |  | 115 |  | Outrage | Dragon | Physical |
|  | 2-2-004 | Charizard (Gigantamax) | 6-Star | Fire | Flying | Gigantamax | 158 | 166 / 174 / 182 / 189 | 61 |  | 166 | 107 |  | 100 |  | 137 |  | 108 |  | 127 |  | Flamethrower | Fire | Special |
|  | 2-2-005 | Zarude | 6-Star | Dark | Grass | Dynamax | 150 | 158 / 165 / 173 / 181 | 61 |  | 199 | 151 |  | 133 |  | 90 |  | 120 |  | 133 |  | Power Whip | Grass | Physical |
|  | 2-2-006 | Lucario | 6-Star | Fighting | Steel | Dynamax | 146 | 154 / 161 / 168 / 176 | 61 |  | 156 | 139 |  | 90 |  | 145 |  | 90 |  | 114 |  | Aura Sphere | Fighting | Special |
|  | 2-2-007 | Rayquaza (Mega) | 6-Star | Dragon | Flying | Mega Evolution | 184 | 202 | 61 |  | 199 | 188 |  | 114 |  | 188 |  | 114 |  | 120 |  | Dragon Ascent | Flying | Physical |
|  | 2-2-008 | Latios | 6-Star | Dragon | Psychic |  | 154 |  | 60 |  | 166 | 101 |  | 113 |  | 137 |  | 161 |  | 137 |  | Mist Ball | Psychic | Special |
|  | 2-2-009 | Latias | 6-Star | Dragon | Psychic |  | 154 |  | 60 |  | 166 | 113 |  | 101 |  | 161 |  | 137 |  | 137 |  | Luster Purge | Psychic | Special |
|  | 2-2-010 | Naganadel | 6-Star | Poison | Dragon |  | 130 |  | 61 |  | 160 | 94 |  | 94 |  | 159 |  | 94 |  | 152 |  | Venoshock | Poison | Special |
|  | 2-2-011 | Metagross (Mega) | 6-Star | Steel | Psychic | Mega Evolution | 144 | 160 | 53 |  | 147 | 148 | 158 | 142 | 164 | 105 | 116 | 100 | 121 | 79 | 121 | Meteor Mash | Steel | Physical |
|  | 2-2-012 | Haxorus |  | Dragon |  | Z-Move | 132 |  | 53 |  | 143 | 160 |  | 100 |  | 68 |  | 79 |  | 107 |  | Outrage | Dragon | Physical |
|  | 2-2-013 | Aerodactyl |  | Rock | Flying | Z-Move | 128 |  | 54 |  | 150 | 118 |  | 75 |  | 69 |  | 86 |  | 145 |  | Fly | Flying | Physical |
|  | 2-2-014 | Dracozolt |  | Electric | Dragon | Dynamax | 138 | 145 / 152 / 159 | 59 |  | 175 | 123 |  | 111 |  | 99 |  | 87 |  | 93 |  | Thunder | Electric | Special |
|  | 2-2-015 | Dracovish |  | Water | Dragon | Dynamax | 138 | 145 / 152 / 159 | 59 |  | 175 | 111 |  | 123 |  | 87 |  | 99 |  | 93 |  | Outrage | Dragon | Physical |
|  | 2-2-016 | Alakazam |  | Psychic |  | Dynamax | 128 | 135 / 141 / 148 | 55 |  | 125 | 60 |  | 54 |  | 153 |  | 109 |  | 137 |  | Psychic | Psychic | Special |
|  | 2-2-017 | Arctozolt |  | Electric | Ice | Tag Move | 130 |  | 56 |  | 172 | 121 |  | 109 |  | 109 |  | 97 |  | 68 |  | Blizzard | Ice | Special |
|  | 2-2-018 | Golisopod |  | Bug | Water | Tag Move | 132 |  | 54 |  | 145 | 140 |  | 156 |  | 69 |  | 102 |  | 48 |  | First Impression | Bug | Physical |
|  | 2-2-019 | Serperior |  | Grass |  | Chain Attack | 144 |  | 59 |  | 157 | 93 |  | 117 |  | 93 |  | 117 |  | 138 |  | Leaf Storm | Grass | Special |
|  | 2-2-020 | Emboar |  | Fire | Fighting | Chain Attack | 144 |  | 59 |  | 198 | 150 |  | 81 |  | 123 |  | 81 |  | 81 |  | Flare Blitz | Fire | Physical |
|  | 2-2-021 | Samurott |  | Water |  | Chain Attack | 144 |  | 59 |  | 181 | 123 |  | 105 |  | 132 |  | 87 |  | 87 |  | Hydro Pump | Water | Special |
|  | 2-2-022 | Iron Treads |  | Ground | Steel |  | 148 |  | 57 |  | 169 | 132 |  | 141 |  | 87 |  | 84 |  | 125 |  | Earthquake | Ground | Physical |
|  | 2-2-023 | Poipole |  | Poison |  |  | 110 |  | 55 |  | 138 | 85 |  | 78 |  | 85 |  | 78 |  | 85 |  | Dragon Pulse | Dragon | Special |
|  | 2-2-024 | Hatterene |  | Psychic | Fairy |  | 128 |  | 54 |  | 125 | 102 |  | 107 |  | 151 |  | 116 |  | 36 |  | Dazzling Gleam | Fairy | Special |
|  | 2-2-025 | Absol |  | Dark |  |  | 120 |  | 60 |  | 148 | 161 |  | 77 |  | 95 |  | 77 |  | 95 |  | Throat Chop | Dark | Physical |
|  | 2-2-026 | Floragato |  | Grass |  |  | 84 |  | 42 |  | 103 | 72 |  | 57 |  | 55 |  | 57 |  | 74 |  | Magical Leaf | Grass | Special |
|  | 2-2-027 | Meowscarada | 4-Star | Grass | Dark |  | 116 |  | 48 |  | 130 | 110 |  | 72 |  | 82 |  | 72 |  | 123 |  | Night Slash | Dark | Physical |
|  | 2-2-028 | Crocalor |  | Fire |  |  | 84 |  | 42 |  | 120 | 51 |  | 70 |  | 80 |  | 53 |  | 46 |  | Incinerate | Fire | Special |
|  | 2-2-029 | Skeledirge | 4-Star | Fire | Ghost |  | 116 |  | 48 |  | 157 | 77 |  | 101 |  | 110 |  | 77 |  | 68 |  | Hex | Ghost | Special |
|  | 2-2-030 | Quaxwell |  | Water |  |  | 84 |  | 42 |  | 110 | 76 |  | 59 |  | 59 |  | 55 |  | 59 |  | Water Pulse | Water | Special |
|  | 2-2-031 | Quaquaval | 4-Star | Water | Fighting |  | 116 |  | 48 |  | 139 | 120 |  | 81 |  | 86 |  | 77 |  | 86 |  | Low Sweep | Fighting | Physical |
|  | 2-2-032 | Snivy |  | Grass |  |  | 46 |  | 27 |  | 61 | 29 |  | 34 |  | 29 |  | 34 |  | 39 |  | Vine Whip | Grass | Physical |
|  | 2-2-033 | Servine |  | Grass |  |  | 84 |  | 41 |  | 102 | 55 |  | 68 |  | 55 |  | 68 |  | 74 |  | Mega Drain | Grass | Special |
|  | 2-2-034 | Serperior | 4-Star | Grass |  |  | 116 |  | 48 |  | 130 | 77 |  | 96 |  | 77 |  | 96 |  | 113 |  | Leaf Tornado | Grass | Special |
|  | 2-2-035 | Tepig |  | Fire |  |  | 46 |  | 27 |  | 72 | 39 |  | 29 |  | 29 |  | 29 |  | 29 |  | Ember | Fire | Special |
|  | 2-2-036 | Pignite |  | Fire | Fighting |  | 80 |  | 41 |  | 127 | 83 |  | 51 |  | 63 |  | 51 |  | 51 |  | Flame Charge | Fire | Physical |
|  | 2-2-037 | Emboar | 4-Star | Fire | Fighting |  | 116 |  | 48 |  | 163 | 123 |  | 67 |  | 101 |  | 67 |  | 67 |  | Fire Punch | Fire | Physical |
|  | 2-2-038 | Oshawott |  | Water |  |  | 46 |  | 27 |  | 66 | 34 |  | 29 |  | 39 |  | 29 |  | 29 |  | Aqua Jet | Water | Physical |
|  | 2-2-039 | Dewott |  | Water |  |  | 84 |  | 41 |  | 115 | 68 |  | 55 |  | 74 |  | 55 |  | 55 |  | Water Pulse | Water | Special |
|  | 2-2-040 | Samurott | 4-Star | Water |  |  | 116 |  | 48 |  | 149 | 101 |  | 86 |  | 108 |  | 72 |  | 72 |  | Razor Shell | Water | Physical |
|  | 2-2-041 | Poliwag |  | Water |  |  | 42 |  | 26 |  | 56 | 31 |  | 25 |  | 25 |  | 25 |  | 51 |  | Water Gun | Water | Special |
|  | 2-2-042 | Poliwhirl |  | Water |  |  | 74 |  | 39 |  | 99 | 55 |  | 55 |  | 44 |  | 44 |  | 75 |  | Bubble Beam | Water | Special |
|  | 2-2-043 | Poliwrath | 4-Star | Water | Fighting |  | 100 |  | 43 |  | 136 | 90 |  | 90 |  | 68 |  | 86 |  | 68 |  | Submission | Fighting | Physical |
|  | 2-2-044 | Feebas |  | Water |  |  | 32 |  | 24 |  | 46 | 12 |  | 15 |  | 10 |  | 33 |  | 46 |  | Tackle | Normal | Physical |
|  | 2-2-045 | Milotic | 4-Star | Water |  |  | 112 |  | 45 |  | 140 | 59 |  | 76 |  | 95 |  | 117 |  | 77 |  | Waterfall | Water | Physical |
|  | 2-2-046 | Tynamo |  | Electric |  |  | 42 |  | 27 |  | 55 | 34 |  | 26 |  | 29 |  | 26 |  | 37 |  | Charge Beam | Electric | Special |
|  | 2-2-047 | Eelektrik |  | Electric |  |  | 82 |  | 41 |  | 104 | 74 |  | 62 |  | 66 |  | 62 |  | 37 |  | Shock Wave | Electric | Special |
|  | 2-2-048 | Eelektross | 4-Star | Electric |  |  | 112 |  | 47 |  | 136 | 113 |  | 80 |  | 103 |  | 80 |  | 52 |  | Volt Switch | Electric | Special |
|  | 2-2-049 | Ralts |  | Psychic | Fairy |  | 34 |  | 27 |  | 52 | 18 |  | 18 |  | 29 |  | 23 |  | 26 |  | Confusion | Psychic | Special |
|  | 2-2-050 | Kirlia |  | Psychic | Fairy |  | 50 |  | 41 |  | 82 | 33 |  | 33 |  | 58 |  | 50 |  | 46 |  | Psybeam | Psychic | Special |
|  | 2-2-051 | Gardevoir | 4-Star | Psychic | Fairy |  | 114 |  | 47 |  | 120 | 66 |  | 66 |  | 122 |  | 113 |  | 80 |  | Psychic | Psychic | Special |
|  | 2-2-052 | Gallade | 4-Star | Psychic | Fighting |  | 114 |  | 47 |  | 120 | 122 |  | 66 |  | 66 |  | 113 |  | 80 |  | Brick Break | Fighting | Physical |
|  | 2-2-053 | Venipede |  | Bug | Poison |  | 40 |  | 27 |  | 53 | 29 |  | 36 |  | 21 |  | 26 |  | 35 |  | Poison Sting | Poison | Physical |
|  | 2-2-054 | Whirlipede |  | Bug | Poison |  | 76 |  | 41 |  | 85 | 51 |  | 55 |  | 38 |  | 71 |  | 44 |  | Bug Bite | Bug | Physical |
|  | 2-2-055 | Scolipede | 4-Star | Bug | Poison |  | 110 |  | 47 |  | 115 | 101 |  | 90 |  | 57 |  | 71 |  | 112 |  | Cross Poison | Poison | Physical |
|  | 2-2-056 | Slakoth |  | Normal |  |  | 42 |  | 26 |  | 72 | 33 |  | 33 |  | 23 |  | 23 |  | 18 |  | Tackle | Normal | Physical |
|  | 2-2-057 | Greedent | 4-Star | Normal |  |  | 98 |  | 43 |  | 163 | 90 |  | 90 |  | 54 |  | 72 |  | 23 |  | Covet | Normal | Physical |
|  | 2-2-058 | Spheal |  | Ice | Water |  | 44 |  | 27 |  | 74 | 26 |  | 32 |  | 34 |  | 32 |  | 18 |  | Powder Snow | Ice | Special |
|  | 2-2-059 | Sealeo |  | Ice | Water |  | 82 |  | 41 |  | 124 | 54 |  | 52 |  | 60 |  | 62 |  | 41 |  | Icy Wind | Ice | Special |
|  | 2-2-060 | Walrein | 4-Star | Ice | Water |  | 110 |  | 47 |  | 160 | 80 |  | 89 |  | 94 |  | 89 |  | 66 |  | Aurora Beam | Ice | Special |
|  | 2-2-061 | Budew |  | Grass | Poison |  | 40 |  | 26 |  | 56 | 20 |  | 23 |  | 31 |  | 41 |  | 33 |  | Absorb | Grass | Special |
|  | 2-2-062 | Roselia |  | Grass | Poison |  | 76 |  | 39 |  | 88 | 51 |  | 40 |  | 83 |  | 67 |  | 55 |  | Mega Drain | Grass | Special |
|  | 2-2-063 | Roserade | 4-Star | Grass | Poison |  | 106 |  | 43 |  | 109 | 68 |  | 63 |  | 117 |  | 99 |  | 86 |  | Magical Leaf | Grass | Special |
|  | 2-2-064 | Milcery |  | Fairy |  |  | 42 |  | 27 |  | 61 | 26 |  | 26 |  | 32 |  | 37 |  | 23 |  | Draining Kiss | Fairy | Special |
|  | 2-2-065 | Alcremie | 4-Star | Fairy |  |  | 112 |  | 43 |  | 120 | 62 |  | 77 |  | 110 |  | 121 |  | 66 |  | Draining Kiss | Fairy | Special |
|  | 2-2-066 | Elekid |  | Electric |  |  | 50 |  | 27 |  | 61 | 39 |  | 24 |  | 40 |  | 34 |  | 56 |  | Thunder Shock | Electric | Special |
|  | 2-2-067 | Electabuzz |  | Electric |  |  | 90 |  | 41 |  | 104 | 73 |  | 51 |  | 82 |  | 74 |  | 91 |  | Shock Wave | Electric | Special |
|  | 2-2-068 | Electivire | 4-Star | Electric |  |  | 118 |  | 47 |  | 127 | 120 |  | 67 |  | 94 |  | 84 |  | 94 |  | Thunder Punch | Electric | Physical |
|  | 2-2-069 | Meowth (Galarian) |  | Steel |  |  | 42 |  | 26 |  | 62 | 38 |  | 33 |  | 25 |  | 25 |  | 25 |  | Metal Claw | Steel | Physical |
|  | 2-2-070 | Perrserker |  | Steel |  |  | 82 |  | 38 |  | 101 | 88 |  | 81 |  | 43 |  | 50 |  | 43 |  | Iron Head | Steel | Physical |
|  | L-3-010 | Venusaur (Mega) | Legacy | Grass | Poison | Mega Evolution | 122 | 126 | 50 |  | 140 | 87 | 105 | 88 | 128 | 105 | 127 | 105 | 125 | 85 |  | Leaf Storm | Grass | Special |
|  | L-3-011 | Charizard (Mega X) | Legacy | Fire | Flying | Mega Evolution | 124 | 124 | 50 |  | 138 | 89 | 135 | 83 | 110 | 114 | 135 | 90 |  | 105 |  | Flamethrower | Fire | Special |
|  | L-3-012 | Blastoise (Mega) | Legacy | Water |  | Mega Evolution | 124 | 126 | 50 |  | 139 | 88 | 108 | 105 | 125 | 90 | 140 | 110 | 120 | 83 |  | Hydro Pump | Water | Special |
|  | L-3-013 | Mew | Legacy | Psychic |  |  | 142 |  | 52 |  | 166 | 109 |  | 109 |  | 109 |  | 109 |  | 109 |  | Psychic | Psychic | Special |
|  | L-3-014 | Zacian (Crowned) | Legacy | Fairy | Steel |  | 162 |  | 50 |  | 152 | 175 |  | 120 |  | 85 |  | 120 |  | 153 |  | Behemoth Blade | Steel | Physical |
|  | L-3-015 | Zamazenta (Crowned) | Legacy | Fighting | Steel |  | 162 |  | 50 |  | 152 | 135 |  | 150 |  | 85 |  | 150 |  | 133 |  | Behemoth Bash | Steel | Physical |
|  | L-3-016 | Reshiram | Legacy | Dragon | Fire |  | 162 |  | 53 |  | 169 | 132 |  | 111 |  | 164 |  | 132 |  | 100 |  | Fusion Flare | Fire | Special |
|  | L-3-017 | Zekrom | Legacy | Dragon | Electric |  | 162 |  | 53 |  | 169 | 164 |  | 132 |  | 132 |  | 111 |  | 100 |  | Fusion Bolt | Electric | Physical |
|  | L-3-018 | Kyurem | Legacy | Dragon | Ice |  | 150 |  | 53 |  | 195 | 142 |  | 100 |  | 142 |  | 100 |  | 105 |  | Blizzard | Ice | Special |
|  | P-1-001 | Celebi | Promo Special | Psychic | Grass |  | 110 |  | 42 |  | 130 | 89 |  | 89 |  | 89 |  | 89 |  | 89 |  | Leaf Storm | Grass | Special |
|  | P-1-002 | Gyarados (Mega) | Promo Special | Water | Flying | Mega Evolution | 134 | 139 | 54 |  | 166 | 140 | 172 | 90 | 122 | 69 | 80 | 113 | 145 | 92 |  | Waterfall | Water | Physical |
|  | P-1-003 | Lapras | Promo Special | Water | Ice |  | 98 |  | 39 |  | 146 | 69 |  | 65 |  | 69 |  | 77 |  | 50 |  | Ice Beam | Ice | Special |
|  | P-1-004 | Zygarde (Complete) | Promo Special | Dragon | Ground |  | 128 |  | 40 |  | 222 | 85 |  | 101 |  | 77 |  | 81 |  | 73 |  | Core Enforcer | Dragon | Special |
|  | P-1-005 | Darkrai | Promo Special | Dark |  |  | 98 |  | 35 |  | 94 | 68 |  | 68 |  | 99 |  | 68 |  | 92 |  | Dark Pulse | Dark | Special |
|  | P (G-002) | Keldeo (Resolute) | Promo Special | Water | Fighting |  | 112 |  | 42 |  | 128 | 65 |  | 80 |  | 113 |  | 80 |  | 95 |  | Secret Sword | Fighting | Special |
|  | P (G-003) | Urshifu (Single Strike) | Promo Special | Fighting | Dark |  | 120 |  | 47 |  | 151 | 127 |  | 99 |  | 64 |  | 61 |  | 96 |  | Wicked Blow | Dark | Physical |
|  | P (G-004) | Kyurem | Promo Special | Dragon | Ice |  | 104 |  | 34 |  | 129 | 93 |  | 66 |  | 93 |  | 66 |  | 69 |  | Blizzard | Ice | Special |
|  | P (G-005) | Urshifu (Rapid Strike) | Promo Special | Fighting | Water |  | 120 |  | 47 |  | 151 | 127 |  | 99 |  | 64 |  | 61 |  | 96 |  | Surging Strikes | Water | Physical |
|  | P (G-006) | Duraludon | Promo Special | Steel | Dragon |  | 124 |  | 50 |  | 130 | 100 |  | 120 |  | 125 |  | 55 |  | 90 |  | Flash Cannon | Steel | Special |
|  | P (G-007) | Celebi | Promo Special | Psychic | Grass |  | 100 |  | 38 |  | 124 | 81 |  | 81 |  | 81 |  | 81 |  | 81 |  | Leaf Storm | Grass | Special |
|  | P (G-008) | Corviknight | Promo Special | Flying | Steel |  | 118 |  | 50 |  | 158 | 92 |  | 110 |  | 58 |  | 90 |  | 72 |  | Brave Bird | Flying | Physical |
|  | P (G-009) | Mew | Promo Special | Psychic |  |  | 135 |  | 45 |  | 139 | 91 |  | 91 |  | 91 |  | 91 |  | 91 |  | Psychic | Psychic | Special |

|  | P (G-021) | Zacian (Crowned) | Promo Special | Fairy | Steel |  | 154 |  | 41 |  | 120 | 144 |  | 99 |  | 70 |  | 99 |  | 120 |  |  |  |  |
|  | P (H-017) | Calyrex (Ice Rider) | Promo Special | Psychic | Ice |  | 124 |  | 40 |  | 130 | 137 |  | 125 |  | 73 |  | 109 |  | 45 |  | Glacial Lance | Ice | Physical |
|  | P (G-012) | Lycanroc (Dusk Form) | Promo Special | Rock |  | Z-Move Signature | 122 |  | 54 |  | 145 | 131 |  | 75 |  | 64 |  | 75 |  | 123 |  | Stone Edge | Rock | Physical |
|  | P (Y-001) | Pikachu | Promo Special | Electric |  | Dynamax | 100 | 107 / 112 / 118 | 53 |  | 117 | 74 |  | 55 |  | 68 |  | 68 |  | 118 |  | Thunderbolt | Electric | Special |
|  | P (Y-002) | Pikachu | Promo Special | Electric |  | Z-Move | 100 |  | 53 |  | 117 | 74 |  | 55 |  | 68 |  | 68 |  | 118 |  | Thunderbolt | Electric | Special |
|  | P (Y-003) | Greninja | Promo Special | Water | Dark | Z-Move | 98 |  | 38 |  | 105 | 79 |  | 57 |  | 85 |  | 60 |  | 100 |  | Water Shuriken | Water | Special |
|  | P (Y-004) | Lucario | Promo Special | Fighting | Steel | Z-Move | 94 |  | 38 |  | 101 | 88 |  | 58 |  | 92 |  | 58 |  | 73 |  | Focus Blast | Fighting | Special |
|  | P (Y-005) | Mimikyu | Promo Special | Ghost | Fairy | Z-Move | 112 |  | 50 | Disguise | 115 | 95 |  | 85 |  | 55 |  | 110 |  | 101 |  | Shadow Claw | Ghost | Physical |
|  | P (Y-006) | Flareon | Promo Special | Fire |  |  | 120 |  | 49 |  | 122 | 132 |  | 63 |  | 98 |  | 112 |  | 68 |  | Flamethrower | Fire | Special |
|  | P (Y-007) | Psyduck | Promo Special | Water |  |  | 100 |  | 74 |  | 158 | 81 |  | 76 |  | 101 |  | 79 |  | 86 |  | Water Gun | Water | Special |
|  | P (Y-008) | Dragonite | Promo Special | Dragon | Flying | Z-Move | 118 |  | 43 |  | 131 | 120 |  | 86 |  | 91 |  | 91 |  | 73 |  | Dragon Rush | Dragon | Physical |
|  | P (Y-009) | Charizard (Gigantamax) | Promo Special | Fire | Flying | Gigantamax | 108 | 135 / 141 / 148 / 155 | 52 |  | 143 | 92 |  | 86 |  | 118 |  | 93 |  | 109 |  | Flamethrower | Fire | Special |
|  | P (Y-010) | Charizard (Gigantamax) | Promo Special | Fire | Flying | Gigantamax | 98 | 103 / 107 / 112 / 118 | 53 |  | 109 | 70 |  | 65 |  | 90 |  | 71 |  | 83 |  | Flamethrower | Fire | Special |
|  | P (Y-011) | Gardevoir (Mega) | Promo Special | Psychic | Fairy | Mega Evolution | 122 | 130 | 51 |  | 130 | 71 | 91 | 71 |  | 132 | 173 | 122 | 142 | 86 | 107 | Psychic | Psychic | Special |
|  | P (Y-012) | Scizor (Mega) | Promo Special | Bug | Steel | Mega Evolution | 128 | 140 | 54 |  | 139 | 145 | 187 | 113 | 150 | 64 | 75 | 92 | 114 | 75 | 86 |  |  |  |
|  | P (Y-013) | Alcremie (Gigantamax) | Promo Special | Fairy |  | Gigantamax | 110 | 116 / 121 / 127 | 40 |  | 118 | 61 |  | 75 |  | 108 |  | 118 |  | 65 |  | Dazzling Gleam | Fairy | Special |
|  | R-1-1 | Pikachu | Regular Tags | Electric |  | Dynamax | 92 | 100 / 105 / 111 | 48 |  | 112 | 71 |  | 53 |  | 65 |  | 65 |  | 113 |  | Thunderbolt | Electric | Special |
|  | R-1-2 | Lucario | Regular Tags | Fighting | Steel | Dynamax | 102 | 108 / 113 / 119 | 42 |  | 110 | 97 |  | 63 |  | 101 |  | 63 |  | 80 |  | Focus Blast | Fighting | Special |
|  | R-1-3 | Snorlax | Regular Tags | Normal |  | Dynamax | 104 | 110 / 115 / 120 | 41 |  | 182 | 95 |  | 58 |  | 58 |  | 95 |  | 29 |  | Giga Impact | Normal | Physical |
|  | R-2-1 | Pikachu (Gigantamax) | Regular Tags | Electric |  | Gigantamax | 102 | 108 / 113 / 119 | 41 |  | 118 | 75 |  | 56 |  | 69 |  | 69 |  | 120 |  | Thunderbolt | Electric | Special |
|  | R-2-2 | Charizard (Gigantamax) | Regular Tags | Fire | Flying | Gigantamax | 126 | 126 / 133 / 139 / 145 | 51 |  | 140 | 90 |  | 84 |  | 116 |  | 91 |  | 107 |  | Flamethrower | Fire | Special |
|  | R-2-3 | Gengar (Gigantamax) | Regular Tags | Ghost | Poison | Gigantamax | 118 | 124 / 130 / 136 | 51 |  | 122 | 71 |  | 66 |  | 137 |  | 81 |  | 117 |  | Shadow Ball | Ghost | Special |
`; // <--- ĐẢM BẢO PHẢI CÓ DẤU NÀY ĐỂ ĐÓNG CHUỖI DATA!
// Đang kiểm tra dữ liệu pokemon này
//|  | P (E-001) | Gyarados (Mega) | Promo Special | Water | Flying | Mega Evolution | 154 | 170 | 54 |  | 166 | 140 | 172 | 90 | 122 | 69 | 80 | 113 | 145 | 92 |  | Waterfall | Water | Physical |

// Hàm phụ để làm sạch chuỗi chỉ số và lấy con số đầu tiên trước dấu gạch chéo / (nếu có)
function cleanStat(value) {
    if (!value) return 0;
    let cleanVal = value.split('/')[0].trim();
    let num = parseInt(cleanVal);
    return isNaN(num) ? 0 : num;
}

// Hàm tự động tạo link hình ảnh theo quy định của các phiên bản
function getPokemonImageUrl(serial) {
    if (!serial) return "https://via.placeholder.com/120x120.png?text=No+Image";
    
    let cleanSerial = serial.trim();
    
    // Nếu Serial Code bắt đầu bằng chữ "R" hoặc "r"
    if (cleanSerial.toUpperCase().startsWith('R')) {
        return `https://webassets-pokemonmezastar.marv.jp/assets/img/ver3/pm_en_${cleanSerial}_f.jpg`;
    }
    
    let parts = cleanSerial.split('-');
    let prefix = parts[0] + "-" + (parts[1] || "");

    switch(prefix) {
        case "1-2": // Bản S2 -> .jpg
            return `https://webassets-pokemonmezastar.marv.jp/assets/img/ver2/pm_en_${cleanSerial}_f.jpg`;
        case "1-3": // Bản S3 -> .jpg
            return `https://webassets-pokemonmezastar.marv.jp/assets/img/ver3/pm_en_${cleanSerial}_f.jpg`;
        case "1-4": // Bản S4 -> .jpg
            return `https://webassets-pokemonmezastar.marv.jp/assets/img/ver4/pm_en_${cleanSerial}_f.jpg`;
        case "2-1": // Bản G1 -> .jpg
            return `https://webassets-pokemonmezastar.marv.jp/assets/publish/20260330110000/pm_en_${cleanSerial}_f.jpg`;
        case "2-2": // Bản G2 -> .png
            return `https://webassets-pokemonmezastar.marv.jp/assets/publish/20260625110000/pm_en_${cleanSerial}_f.png`;
        default:
            // Mặc định cho bản S1 cũ hoặc dự phòng
            return `https://webassets-pokemonmezastar.marv.jp/assets/img/ver1/pm_en_${cleanSerial}_f.png`;
    }
}

// Bộ phân tích cú pháp (Parser) xử lý dữ liệu thô
const pokemonData = [];
const lines = rawData.trim().split('\n');

for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    
    if (!line.startsWith('|') || line.includes('Serial Code') || line.includes('---')) continue;
    
    let cols = line.split('|').map(c => c.trim());
    
    if (cols.length >= 25 && cols[2]) {
        let serial = cols[2];
        let weight = 1; 
        let prefix = serial.substring(0, 3);
        
        if (prefix === "1-1") weight = 1;
        else if (prefix === "1-2") weight = 2;
        else if (prefix === "1-3") weight = 3;
        else if (prefix === "1-4") weight = 4;
        else if (prefix === "2-1") weight = 5;
        else if (prefix === "2-2") weight = 6;

        // Ép kiểu chỉ số an toàn bằng hàm cleanStat
        let hp   = cleanStat(cols[12]);
        let atk  = cleanStat(cols[14]) || cleanStat(cols[13]) || 0;
        let def  = cleanStat(cols[16]) || cleanStat(cols[15]) || 0;
        let spa  = cleanStat(cols[18]) || cleanStat(cols[17]) || 0;
        let spd  = cleanStat(cols[20]) || cleanStat(cols[19]) || 0;
        let spe  = cleanStat(cols[22]) || cleanStat(cols[21]) || 0;

        pokemonData.push({
            serialCode: serial,
            name: cols[3],
            rarity: cols[4],
            type1: cols[5],
            type2: cols[6] || "",
            gimmick: cols[7] || "",
            hp: hp,
            atk: atk,
            def: def,
            spa: spa,
            spd: spd,
            spe: spe,
            moveName: cols[23] || "",
            moveType: cols[24] || "",
            moveCategory: cols[25] || "Physical",
            image: getPokemonImageUrl(serial),
            weight: weight,
			isLegacy: cols[4] ? cols[4].toUpperCase().includes('LEGACY') : false
        });
    }
}