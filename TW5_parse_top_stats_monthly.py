#!/usr/bin/env python3

#    parse_top_stats_detailed.py outputs detailed top stats in arcdps logs as parsed by Elite Insights.
#    Copyright (C) 2021 Freya Fleckenstein
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.


import argparse
import datetime
import os.path
from os import listdir
import sys
import xml.etree.ElementTree as ET
from enum import Enum
import importlib
import xlwt

from collections import OrderedDict
from TW5_parse_top_stats_tools import *

if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='This reads a set of arcdps reports in xml format and generates top stats.')
	parser.add_argument('input_directory', help='Directory containing .xml or .json files from arcdps reports')
	parser.add_argument('-o', '--output', dest="output_filename", help="Text file to write the computed top stats")
	#parser.add_argument('-f', '--input_filetype', dest="filetype", help="filetype of input files. Currently supports json and xml, defaults to json.", default="json")
	parser.add_argument('-x', '--xls_output', dest="xls_output_filename", help="xls file to write the computed top stats")    
	parser.add_argument('-j', '--json_output', dest="json_output_filename", help="json file to write the computed top stats to")    
	parser.add_argument('-l', '--log_file', dest="log_file", help="Logging file with all the output")
	parser.add_argument('-c', '--config_file', dest="config_file", help="Config file with all the settings", default="TW5_parser_config_detailed")
	parser.add_argument('-a', '--anonymized', dest="anonymize", help="Create an anonymized version of the top stats. All account and character names will be replaced.", default=False, action='store_true')
	args = parser.parse_args()

	if not os.path.isdir(args.input_directory):
		print("Directory ",args.input_directory," is not a directory or does not exist!")
		sys.exit()
	if args.output_filename is None:
		args.output_filename = args.input_directory+"/TW5_top_stats_detailed.tid"
	if args.xls_output_filename is None:
		args.xls_output_filename = args.input_directory+"/TW5_top_stats_detailed.xls"
	if args.json_output_filename is None:
		args.json_output_filename = args.input_directory+"/TW5_top_stats_detailed.json"                
	if args.log_file is None:
		args.log_file = args.input_directory+"/log_detailed.txt"

	output = open(args.output_filename, "w",encoding="utf-8")
	log = open(args.log_file, "w")

	parser_config = importlib.import_module("parser_configs."+args.config_file , package=None) 
	
	config = fill_config(parser_config)

	print_string = "Using input directory "+args.input_directory+", writing output to "+args.output_filename+" and log to "+args.log_file
	print(print_string)
	print_string = "Considering fights with at least "+str(config.min_allied_players)+" allied players and at least "+str(config.min_enemy_players)+" enemies that took longer than "+str(config.min_fight_duration)+" s."
	myprint(log, print_string)

	players, fights, found_healing, found_barrier, squad_comp, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, auras_TableIn, auras_TableOut, Death_OnTag = collect_stat_data(args, config, log, args.anonymize)    

	# create xls file if it doesn't exist
	book = xlwt.Workbook(encoding="utf-8")
	book.add_sheet("fights overview")
	book.save(args.xls_output_filename)

	
	#Create Tid file header to support drag and drop onto html page
	myDate = datetime.datetime.now()

	myprint(output, 'created: '+myDate.strftime("%Y%m%d%H%M%S"))
	myprint(output, 'creator: Drevarr ')
	myprint(output, 'caption: '+myDate.strftime("%Y%m%d")+'-WvW-Log-Review')
	myprint(output, 'curTab: Overview')
	myprint(output, 'curFight: Fight-1')
	myprint(output, 'curControl-In: Blinded')
	myprint(output, 'curControl-Out: Blinded')
	myprint(output, 'curAuras-Out: Fire')
	myprint(output, 'curAuras-In: Fire')
	myprint(output, 'tags: Logs [['+myDate.strftime("%Y")+'-'+myDate.strftime("%m")+' Log Reviews]]')
	myprint(output, 'title: '+myDate.strftime("%Y%m%d")+'-WvW-Log-Review\n')
	#End Tid file header

	#JEL-Tweaked to output TW5 formatting (https://drevarr.github.io/FluxCapacity.html)
	print_string = "__''Flux Capacity Node Farmers - WVW Log Review''__\n"
	myprint(output, print_string)

	# print overall stats
	overall_squad_stats = get_overall_squad_stats(fights, config)
	overall_raid_stats = get_overall_raid_stats(fights)
	total_fight_duration = print_total_squad_stats(fights, overall_squad_stats, overall_raid_stats, found_healing, found_barrier, config, output)

	#Start nav_bar_menu for TW5
	Nav_Bar_Items= ('<$button set="!!curTab" setTo="Overview" selectedClass="" class="btn btn-sm btn-dark" style=""> Session Overview </$button>',
					'<$button set="!!curTab" setTo="Deaths" selectedClass="" class="btn btn-sm btn-dark" style=""> Deaths </$button>',
					'<$button set="!!curTab" setTo="Illusion of Life" selectedClass="" class="btn btn-sm btn-dark" style=""> IOL </$button>',
					'<$button set="!!curTab" setTo="Resurrect" selectedClass="" class="btn btn-sm btn-dark" style=""> Resurrect </$button>',                    
					'<$button set="!!curTab" setTo="Enemies Downed" selectedClass="" class="btn btn-sm btn-dark" style=""> Enemies Downed </$button>',
					'<$button set="!!curTab" setTo="Enemies Killed" selectedClass="" class="btn btn-sm btn-dark" style=""> Enemies Killed </$button>',
					'<$button set="!!curTab" setTo="Damage" selectedClass="" class="btn btn-sm btn-dark" style=""> Damage </$button>',
					'<$button set="!!curTab" setTo="Power Damage" selectedClass="" class="btn btn-sm btn-dark" style=""> Power Damage </$button>',
					'<$button set="!!curTab" setTo="Condi Damage" selectedClass="" class="btn btn-sm btn-dark" style=""> Condi Damage </$button>',
					'<$button set="!!curTab" setTo="Damage Taken" selectedClass="" class="btn btn-sm btn-dark" style=""> Damage Taken</$button>',
					'<$button set="!!curTab" setTo="Boon Strips" selectedClass="" class="btn btn-sm btn-dark" style=""> Boon Strips </$button>',
					'<$button set="!!curTab" setTo="Condition Cleanses" selectedClass="" class="btn btn-sm btn-dark" style=""> Condition Cleanses</$button>',
					'<$button set="!!curTab" setTo="Superspeed" selectedClass="" class="btn btn-sm btn-dark" style=""> Superspeed </$button>',
					'<$button set="!!curTab" setTo="Stealth" selectedClass="" class="btn btn-sm btn-dark" style=""> Stealth </$button>',
					'<$button set="!!curTab" setTo="Hide in Shadows" selectedClass="" class="btn btn-sm btn-dark" style=""> Hide in Shadows </$button>',
					'<$button set="!!curTab" setTo="Distance to Tag" selectedClass="" class="btn btn-sm btn-dark" style=""> Distance to Tag </$button>',
					'<$button set="!!curTab" setTo="Stability" selectedClass="" class="btn btn-sm btn-dark" style=""> Stability </$button>',
					'<$button set="!!curTab" setTo="Protection" selectedClass="" class="btn btn-sm btn-dark" style=""> Protection </$button>',
					'<$button set="!!curTab" setTo="Aegis" selectedClass="" class="btn btn-sm btn-dark" style=""> Aegis </$button>',
					'<$button set="!!curTab" setTo="Might" selectedClass="" class="btn btn-sm btn-dark" style=""> Might </$button>',
					'<$button set="!!curTab" setTo="Fury" selectedClass="" class="btn btn-sm btn-dark" style=""> Fury </$button>',
					'<$button set="!!curTab" setTo="Resistance" selectedClass="" class="btn btn-sm btn-dark" style=""> Resistance </$button>',
					'<$button set="!!curTab" setTo="Resolution" selectedClass="" class="btn btn-sm btn-dark" style=""> Resolution </$button>',
					'<$button set="!!curTab" setTo="Quickness" selectedClass="" class="btn btn-sm btn-dark" style=""> Quickness </$button>',
					'<$button set="!!curTab" setTo="Swiftness" selectedClass="" class="btn btn-sm btn-dark" style=""> Swiftness </$button>',
					'<$button set="!!curTab" setTo="Alacrity" selectedClass="" class="btn btn-sm btn-dark" style=""> Alacrity </$button>',
					'<$button set="!!curTab" setTo="Vigor" selectedClass="" class="btn btn-sm btn-dark" style=""> Vigor </$button>',
					'<$button set="!!curTab" setTo="Regeneration" selectedClass="" class="btn btn-sm btn-dark" style=""> Regeneration </$button>',
					'<$button set="!!curTab" setTo="Support" selectedClass="" class="btn btn-sm btn-dark" style=""> Support Players </$button>',
					'<$button set="!!curTab" setTo="Healing" selectedClass="" class="btn btn-sm btn-dark" style=""> Healing </$button>',
					'<$button set="!!curTab" setTo="Barrier" selectedClass="" class="btn btn-sm btn-dark" style=""> Barrier </$button>',
					'<$button set="!!curTab" setTo="Weapon Swaps" selectedClass="" class="btn btn-sm btn-dark" style=""> Weapon Swaps </$button>',
					'<$button set="!!curTab" setTo="Control Effects - Out" selectedClass="" class="btn btn-sm btn-dark" style=""> Control Effects Outgoing </$button>',
					'<$button set="!!curTab" setTo="Control Effects - In" selectedClass="" class="btn btn-sm btn-dark" style=""> Control Effects Incoming </$button>',					
					'<$button set="!!curTab" setTo="Spike Damage" selectedClass="" class="btn btn-sm btn-dark" style=""> Spike Damage </$button>',
					'<$button set="!!curTab" setTo="Buff Uptime" selectedClass="" class="btn btn-sm btn-dark" style=""> Buff Uptime </$button>',
					'<$button set="!!curTab" setTo="Auras - In" selectedClass="" class="btn btn-sm btn-dark" style=""> Auras - In </$button>',
					'<$button set="!!curTab" setTo="Auras - Out" selectedClass="" class="btn btn-sm btn-dark" style=""> Auras - Out </$button>',
					'<$button set="!!curTab" setTo="Death_OnTag" selectedClass="" class="btn btn-sm btn-dark" style=""> Death OnTag </$button>',
					'<$button set="!!curTab" setTo="Downed_Healing" selectedClass="" class="btn btn-sm btn-dark" style=""> Downed Healing </$button>',
					'<$button set="!!curTab" setTo="Offensive Stats" selectedClass="" class="btn btn-sm btn-dark" style=""> Offensive Stats </$button>'
	)
	for item in Nav_Bar_Items:
		myprint(output, item)
	
	myprint(output, '\n---\n')

	#End nav_bar_menu for TW5

	#Overview reveal
	myprint(output, '<$reveal type="match" state="!!curTab" text="Overview">')
	myprint(output, '\n!!!OVERVIEW\n')

	print_fights_overview(fights, overall_squad_stats, overall_raid_stats, config, output)

	#End reveal
	myprint(output, '</$reveal>')

	write_fights_overview_xls(fights, overall_squad_stats, overall_raid_stats, config, args.xls_output_filename)
	
	# print top x players for all stats. If less then x
	# players, print all. If x-th place doubled, print all with the
	# same amount of top x achieved.
	num_used_fights = overall_raid_stats['num_used_fights']

	top_total_stat_players = {key: list() for key in config.stats_to_compute}
	top_consistent_stat_players = {key: list() for key in config.stats_to_compute}
	top_average_stat_players = {key: list() for key in config.stats_to_compute}
	top_percentage_stat_players = {key: list() for key in config.stats_to_compute}
	top_late_players = {key: list() for key in config.stats_to_compute}
	top_jack_of_all_trades_players = {key: list() for key in config.stats_to_compute}    
	
	#JEL-Tweaked to output TW5 formatting (https://drevarr.github.io/FluxCapacity.html)
	for stat in config.stats_to_compute:
		if (stat == 'heal' and not found_healing) or (stat == 'barrier' and not found_barrier):
			continue
		
		fileDate = myDate

		#JEL-Tweaked to output TW5 output to maintain formatted table and slider (https://drevarr.github.io/FluxCapacity.html)
		myprint(output,'<$reveal type="match" state="!!curTab" text="'+config.stat_names[stat]+'">')
		myprint(output, "\n!!!"+config.stat_names[stat].upper()+"\n")

		if stat == 'dist':
			myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			top_consistent_stat_players[stat] = get_top_players(players, config, stat, StatType.CONSISTENT)
			top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
			top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)            
			top_percentage_stat_players[stat],comparison_val = get_and_write_sorted_top_percentage(players, config, num_used_fights, stat, output, StatType.PERCENTAGE, top_consistent_stat_players[stat])
			myprint(output, '\n</div>\n    <div class="flex-col border">\n')
			top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
			top_average_stat_players[stat] = get_and_write_sorted_average(players, config, num_used_fights, stat, output)
			myprint(output, '\n</div>\n</div>\n')
		elif stat == 'dmg_taken':
			myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			top_consistent_stat_players[stat] = get_top_players(players, config, stat, StatType.CONSISTENT)
			top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
			top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
			top_average_stat_players[stat] = get_and_write_sorted_average(players, config, num_used_fights, stat, output)
			myprint(output, '\n</div>\n</div>\n')
		else:
			myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			top_consistent_stat_players[stat] = get_and_write_sorted_top_consistent(players, config, num_used_fights, stat, output)
			myprint(output, '\n</div>\n    <div class="flex-col border">\n')
			top_total_stat_players[stat] = get_and_write_sorted_total(players, config, total_fight_duration, stat, output)
			myprint(output, '\n</div>\n</div>\n')
			top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)
			top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
			
			myprint(output, '<div>')
			myprint(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData}} $height="600px" $theme="dark"/>')
			myprint(output, '</div>')
		#JEL-Tweaked to output TW5 output to maintain formatted table and slider (https://drevarr.github.io/FluxCapacity.html)
		myprint(output, "</$reveal>\n")

		write_to_json(overall_raid_stats, overall_squad_stats, fights, players, top_total_stat_players, top_average_stat_players, top_consistent_stat_players, top_percentage_stat_players, top_late_players, top_jack_of_all_trades_players, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, auras_TableIn, auras_TableOut, Death_OnTag, args.json_output_filename)

	#print table of accounts that fielded support characters
	myprint(output,'<$reveal type="match" state="!!curTab" text="Support">')
	myprint(output, "\n")
	# print table header
	print_string = "|thead-dark table-hover sortable|k"    
	myprint(output, print_string)
	print_string = "|!Account |!Name |!Profession | !Fights| !Duration|!Support |!Guild Status |h"
	myprint(output, print_string)    

	for stat in config.stats_to_compute:
		if (stat == 'rips' or stat == 'cleanses' or stat == 'stability'):
			write_support_players(players, top_total_stat_players[stat], stat, output)

	myprint(output, "</$reveal>\n")

	supportCount=0

	#start Control Effects Outgoing insert
	myprint(output, '<$reveal type="match" state="!!curTab" text="Control Effects - Out">')    
	myprint(output, '\n<<alert-leftbar success "Outgoing Control Effects generated by the Squad" width:60%, class:"font-weight-bold">>\n\n')
	Control_Effects = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark"}
	for C_E in Control_Effects:
		myprint(output, '<$button set="!!curControl-Out" setTo="'+Control_Effects[C_E]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Control_Effects[C_E]+' </$button>')
	
	myprint(output, '\n---\n')
	

	for C_E in Control_Effects:
		key = Control_Effects[C_E]
		if key in squad_Control:
			sorted_squadControl = dict(sorted(squad_Control[key].items(), key=lambda x: x[1], reverse=True))

			i=1
		
			myprint(output, '<$reveal type="match" state="!!curControl-Out" text="'+key+'">\n')
			myprint(output, '\n---\n')
			myprint(output, "|table-caption-top|k")
			myprint(output, "|{{"+key+"}} "+key+" output by Squad Player Descending [TOP 25 Max]|c")
			myprint(output, "|thead-dark table-hover|k")
			myprint(output, "|Place |Name | Profession | Total|h")
			
			for name in sorted_squadControl:
				prof = "Not Found"
				counter = 0
				for nameIndex in players:
					if nameIndex.name == name:
						prof = nameIndex.profession

				if i <=25:
					myprint(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_squadControl[name], 4))+"|")
					i=i+1

			myprint(output, "</$reveal>\n")

			write_control_effects_out_xls(sorted_squadControl, key, players, args.xls_output_filename)
	myprint(output, "</$reveal>\n")
	#end Control Effects Outgoing insert

	#start Control Effects Incoming insert
	myprint(output, '<$reveal type="match" state="!!curTab" text="Control Effects - In">')    
	myprint(output, '\n<<alert-leftbar danger "Incoming Control Effects generated by the Enemy" width:60%, class:"font-weight-bold">>\n\n')
	Control_Effects = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark"}
	for C_E in Control_Effects:
		myprint(output, '<$button set="!!curControl-In" setTo="'+Control_Effects[C_E]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Control_Effects[C_E]+' </$button>')
	
	myprint(output, '\n---\n')
	

	for C_E in Control_Effects:
		key = Control_Effects[C_E]
		if key in enemy_Control:
			sorted_enemyControl = dict(sorted(enemy_Control[key].items(), key=lambda x: x[1], reverse=True))

			i=1
			
			myprint(output, '<$reveal type="match" state="!!curControl-In" text="'+key+'">\n')
			myprint(output, '\n---\n')
			myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			myprint(output, "|table-caption-top|k")
			myprint(output, "|{{"+key+"}} "+key+" impacted Squad Player Descending [TOP 25 Max]|c")
			myprint(output, "|thead-dark table-hover|k")
			myprint(output, "|Place |Name | Profession | Total|h")
			
			for name in sorted_enemyControl:
				prof = "Not Found"
				counter = 0
				for nameIndex in players:
					if nameIndex.name == name:
						prof = nameIndex.profession

				if i <=25:
					myprint(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_enemyControl[name], 4))+"|")
					i=i+1

			#myprint(output, "</$reveal>\n")

			write_control_effects_in_xls(sorted_enemyControl, key, players, args.xls_output_filename)

		if key in enemy_Control_Player:
			sorted_enemyControlPlayer = dict(sorted(enemy_Control_Player[key].items(), key=lambda x: x[1], reverse=True))

			i=1
	
			myprint(output, '\n---\n')
			myprint(output, '\n</div>\n    <div class="flex-col border">\n')
			myprint(output, "|table-caption-top|k")
			myprint(output, "|{{"+key+"}} "+key+" output by Enemy Player Descending [TOP 25 Max]|c")
			myprint(output, "|thead-dark table-hover|k")
			myprint(output, "|Place |Name | Profession | Total|h")
		
			for name in sorted_enemyControlPlayer:
				prof = name.split(' pl')[0]
				counter = 0

				if i <=25:
					myprint(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_enemyControlPlayer[name],4 ))+"|")
					i=i+1

			myprint(output, '\n</div>\n</div>\n')
			myprint(output, "</$reveal>\n")

	myprint(output, "</$reveal>\n")
	#end Control Effects Incoming insert

	#start Aura Effects Incoming insert
	myprint(output, '<$reveal type="match" state="!!curTab" text="Auras - In">')    
	myprint(output, '\n<<alert-leftbar danger "Auras by receiving Player" width:60%, class:"font-weight-bold">>\n\n')
	Auras_Order = {5677: 'Fire', 5577: 'Shocking', 5579: 'Frost', 5684: 'Magnetic', 25518: 'Light', 39978: 'Dark', 10332: 'Chaos'}
	for Aura in Auras_Order:
		myprint(output, '<$button set="!!curAuras-In" setTo="'+Auras_Order[Aura]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Auras_Order[Aura]+' Aura </$button>')
	
	myprint(output, '\n---\n')
	

	for Aura in Auras_Order:
		key = Auras_Order[Aura]
		if key in auras_TableIn:
			sorted_auras_TableIn = dict(sorted(auras_TableIn[key].items(), key=lambda x: x[1], reverse=True))

			i=1
		
			myprint(output, '<$reveal type="match" state="!!curAuras-In" text="'+key+'">\n')
			myprint(output, '\n---\n')
			myprint(output, "|table-caption-top|k")
			myprint(output, "|{{"+key+"}} "+key+" Aura received by Squad Player Descending [TOP 25 Max]|c")
			myprint(output, "|thead-dark table-hover|k")
			myprint(output, "|Place |Name | Profession | Total|h")
			
			for name in sorted_auras_TableIn:
				prof = "Not Found"
				counter = 0
				for nameIndex in players:
					if nameIndex.name == name:
						prof = nameIndex.profession

				if i <=25:
					myprint(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_auras_TableIn[name], 4))+"|")
					i=i+1

			myprint(output, "</$reveal>\n")

			write_auras_in_xls(sorted_auras_TableIn, key, players, args.xls_output_filename)
	myprint(output, "</$reveal>\n")
	#end Auras Incoming insert

	#start Aura Effects Out insert
	myprint(output, '<$reveal type="match" state="!!curTab" text="Auras - Out">')    
	myprint(output, '\n<<alert-leftbar info "Auras output by Player" width:60%, class:"font-weight-bold">>\n\n')
	Auras_Order = {5677: 'Fire', 5577: 'Shocking', 5579: 'Frost', 5684: 'Magnetic', 25518: 'Light', 39978: 'Dark', 10332: 'Chaos'}
	for Aura in Auras_Order:
		myprint(output, '<$button set="!!curAuras-Out" setTo="'+Auras_Order[Aura]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Auras_Order[Aura]+' Aura </$button>')
	
	myprint(output, '\n---\n')
	

	for Aura in Auras_Order:
		key = Auras_Order[Aura]
		if key in auras_TableOut:
			sorted_auras_TableOut = dict(sorted(auras_TableOut[key].items(), key=lambda x: x[1], reverse=True))

			i=1
		
			myprint(output, '<$reveal type="match" state="!!curAuras-Out" text="'+key+'">\n')
			myprint(output, '\n---\n')
			myprint(output, "|table-caption-top|k")
			myprint(output, "|{{"+key+"}} "+key+" Aura output by Squad Player Descending [TOP 10 Max]|c")
			myprint(output, "|thead-dark table-hover|k")
			myprint(output, "|Place |Name | Profession | Total|h")
			
			for name in sorted_auras_TableOut:
				prof = "Not Found"
				counter = 0
				for nameIndex in players:
					if nameIndex.name == name:
						prof = nameIndex.profession

				if i <=10:
					myprint(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_auras_TableOut[name], 4))+"|")
					i=i+1

			myprint(output, "</$reveal>\n")

			write_auras_out_xls(sorted_auras_TableOut, key, players, args.xls_output_filename)
	myprint(output, "</$reveal>\n")
	#end Auras Out insert

	#start Buff Uptime Table insert
	uptime_Order = ['stability',  'protection',  'aegis',  'might',  'fury',  'resistance',  'resolution',  'quickness',  'swiftness',  'alacrity',  'vigor',  'regeneration']
	myprint(output, '<$reveal type="match" state="!!curTab" text="Buff Uptime">')    
	myprint(output, '\n<<alert-leftbar light "Total Buff Uptime % across all fights attended.\n Current Formula: (((Sum of (fight duration * Uptime%))/Attendance)*100)" width:60%, class:"font-weight-bold">>\n\n')
	
	myprint(output, '\n---\n')
	myprint(output, '\n---\n')

	myprint(output, "|table-caption-top|k")
	myprint(output, "|Sortable table - Click header item to sort table |c")
	myprint(output, "|thead-dark table-hover sortable|k")
	myprint(output, "|!Name | !Profession | !Attendance| !{{Stability}}|  !{{Protection}}|  !{{Aegis}}|  !{{Might}}|  !{{Fury}}|  !{{Resistance}}|  !{{Resolution}}|  !{{Quickness}}|  !{{Swiftness}}|  !{{Alacrity}}|  !{{Vigor}}|  !{{Regeneration}}|h")
	for squadDps_prof_name in uptime_Table:
		fightTime = uptime_Table[squadDps_prof_name]['duration']
		name = uptime_Table[squadDps_prof_name]['name']
		prof = uptime_Table[squadDps_prof_name]['prof']

		output_string = "|"+name+" |"
		output_string += " {{"+prof+"}} | "+str(fightTime)+"|"
		for item in uptime_Order:
			if item in uptime_Table[squadDps_prof_name]:
				output_string += " "+str(round(((uptime_Table[squadDps_prof_name][item]/fightTime)*100), 4))+"|"
			else:
				output_string += " 0.00|"
				


		myprint(output, output_string)

	write_buff_uptimes_in_xls(uptime_Table, players, uptime_Order, args.xls_output_filename)
	myprint(output, "</$reveal>\n")
	#end Buff Uptime Table insert

	#start On Tag Death insert
	myprint(output, '<$reveal type="match" state="!!curTab" text="Death_OnTag">')    
	myprint(output, '\n<<alert-leftbar light "On Tag Death Review \n Current Formula: (On Tag <= 600 Range, Off Tag >600 and <=5000, Run Back Death > 5000)" width:60%, class:"font-weight-bold">>\n\n')
	
	myprint(output, '\n---\n')
	myprint(output, '\n---\n')

	myprint(output, "|table-caption-top|k")
	myprint(output, "|Sortable table - Click header item to sort table |c")
	myprint(output, "|thead-dark table-hover sortable|k")
	myprint(output, "|!Name | !Profession | !Attendance | !On_Tag |  !Off_Tag | !After_Tag_Death |  !Run_Back |  !Total | Off Tag Ranges |h")
	for deathOnTag_prof_name in Death_OnTag:
		name = Death_OnTag[deathOnTag_prof_name]['name']
		prof = Death_OnTag[deathOnTag_prof_name]['profession']
		fightTime = uptime_Table[deathOnTag_prof_name]['duration']

		if Death_OnTag[deathOnTag_prof_name]['Off_Tag']:
			converted_list = [str(round(element)) for element in Death_OnTag[deathOnTag_prof_name]['Ranges']]
			Ranges_string = ",".join(converted_list)
		else:
			Ranges_string = " "

		output_string = "|"+name+" |"
		output_string += " {{"+prof+"}} | "+str(fightTime)+" | "+str(Death_OnTag[deathOnTag_prof_name]['On_Tag'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['Off_Tag'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['After_Tag_Death'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['Run_Back'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['Total'])+" |"+Ranges_string+" |"
	


		myprint(output, output_string)

	write_Death_OnTag_xls(Death_OnTag, uptime_Table, players, args.xls_output_filename)
	myprint(output, "</$reveal>\n")
	#end On Tag Death insert

	#Downed Healing
	down_Heal_Order = {14419: 'Battle Standard', 9163: 'Signet of Mercy', 5763: 'Renewal of Water', 5762: 'Renewal of Fire', 5760: 'Renewal of Air', 5761: 'Renewal of Earth'}
	myprint(output, '<$reveal type="match" state="!!curTab" text="Downed_Healing">')    
	myprint(output, '\n<<alert-leftbar light "Healing to downed players (Instant Revive Skills) - requires Heal Stat addon for ARCDPS to track" width:60%, class:"font-weight-bold">>\n\n')
	
	myprint(output, '\n---\n')
	myprint(output, '\n---\n')

	myprint(output, '\n<div class="flex-row">\n<div class="flex-col border">\n')
	myprint(output, "\n!!Healing done\nWork in Progress more skills to be added when logs available\n")
	myprint(output, "|table-caption-top|k")
	myprint(output, "|Sortable table - Click header item to sort table |c")
	myprint(output, "|thead-dark table-hover sortable|k")
	output_string = "|!Name | !Profession | !Attendance |"
	for item in down_Heal_Order:
		output_string += "!{{"+down_Heal_Order[item]+"}}|"
	output_string += "h"
	myprint(output, output_string)
	for squadDps_prof_name in downed_Healing:
		name = downed_Healing[squadDps_prof_name]['name']
		prof = downed_Healing[squadDps_prof_name]['prof']
		fightTime = uptime_Table[squadDps_prof_name]['duration']

		output_string = "|"+name+" |{{"+prof+"}}|"+str(fightTime)+"| "
		for skill in down_Heal_Order:
			if down_Heal_Order[skill] in downed_Healing[squadDps_prof_name]:
				output_string += str(downed_Healing[squadDps_prof_name][down_Heal_Order[skill]]['Heals'])+"|"
			else:
				output_string += " |"
		myprint(output, output_string)
	
	myprint(output, '\n</div>\n<div class="flex-col border">\n')
	myprint(output, "\n!!Number of Skill Hits\nWork in Progress more skills to be added when logs available\n")
	myprint(output, "|table-caption-top|k")
	myprint(output, "|Sortable table - Click header item to sort table |c")
	myprint(output, "|thead-dark table-hover sortable|k")
	output_string = "|!Name | !Profession | !Attendance |"
	for item in down_Heal_Order:
		output_string += "!{{"+down_Heal_Order[item]+"}}|"
	output_string += "h"
	myprint(output, output_string)
	for squadDps_prof_name in downed_Healing:
		name = downed_Healing[squadDps_prof_name]['name']
		prof = downed_Healing[squadDps_prof_name]['prof']
		fightTime = uptime_Table[squadDps_prof_name]['duration']

		output_string = "|"+name+" |{{"+prof+"}}|"+str(fightTime)+"| "
		for skill in down_Heal_Order:
			if down_Heal_Order[skill] in downed_Healing[squadDps_prof_name]:
				output_string += str(downed_Healing[squadDps_prof_name][down_Heal_Order[skill]]['Hits'])+" |"
			else:
				output_string += " |"
		myprint(output, output_string)



	myprint(output, '\n</div>\n</div>\n</$reveal>\n')
	#End Downed Healing

	#start Offensive Stat Table insert
	offensive_Order = ['Critical',  'Flanking',  'Glancing',  'Moving',  'Blinded',  'Interupt',  'Invulnerable',  'Evaded',  'Blocked']
	myprint(output, '<$reveal type="match" state="!!curTab" text="Offensive Stats">')    
	myprint(output, '\n<<alert-leftbar light "Offensive Stats across all fights attended." width:60%, class:"font-weight-bold">>\n\n')
	
	myprint(output, '\n---\n')
	myprint(output, '\n---\n')

	myprint(output, "|table-caption-top|k")
	myprint(output, "|Sortable table - Click header item to sort table |c")
	myprint(output, "|thead-dark table-hover sortable|k")
	myprint(output, "|!Name | !Profession | !{{Critical}}% |  !{{Flanking}}% |  !{{Glancing}}% |  !{{Moving}}% |  !{{Blind}} |  !{{Interupt}} |  !{{Invulnerable}} |  !{{Evaded}} |  !{{Blocked}} |h")
	for squadDps_prof_name in squad_offensive:
		name = squad_offensive[squadDps_prof_name]['name']
		prof = squad_offensive[squadDps_prof_name]['prof']

		output_string = "|"+name+" | {{"+prof+"}} | "

		#Calculate Critical_Hits_Rate
		if squad_offensive[squadDps_prof_name]['stats']['criticalRate']:
			Critical_Rate = round((squad_offensive[squadDps_prof_name]['stats']['criticalRate']/squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount'])*100, 4)
		else:
			Critical_Rate = 0.0000
		Critical_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['criticalRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount'])+' critable hits">'+str(Critical_Rate)+'</span>'
		
		output_string += str(Critical_Rate_TT)+" | "
		
		#Calculate Flanking_Rate
		if squad_offensive[squadDps_prof_name]['stats']['flankingRate']:
			Flanking_Rate = round((squad_offensive[squadDps_prof_name]['stats']['flankingRate']/squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])*100, 4)
		else:
			Flanking_Rate = 0.0000
		Flanking_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['flankingRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' connected direct hit(s)">'+str(Flanking_Rate)+'</span>'
		
		output_string += str(Flanking_Rate_TT)+" | "
		
		#Calculate Glancing Rate
		if squad_offensive[squadDps_prof_name]['stats']['glanceRate']:
			Glancing_Rate = round((squad_offensive[squadDps_prof_name]['stats']['glanceRate']/squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])*100, 4)
		else:
			Glancing_Rate = 0.0000
		Glancing_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['glanceRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' connected direct hit(s)">'+str(Glancing_Rate)+'</span>'
		
		output_string += str(Glancing_Rate_TT)+" | "
		
		#Calculate Moving_Rate
		if squad_offensive[squadDps_prof_name]['stats']['againstMovingRate']:
			Moving_Rate = round((squad_offensive[squadDps_prof_name]['stats']['againstMovingRate']/squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])*100, 4)
		else:
			Moving_Rate = 0.0000
		Moving_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['againstMovingRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])+' direct hit(s)">'+str(Moving_Rate)+'</span>'
		
		output_string += str(Moving_Rate_TT)+" | "
		
		#Calculate Blinded_Rate
		if squad_offensive[squadDps_prof_name]['stats']['missed']:
			Blinded_Rate = squad_offensive[squadDps_prof_name]['stats']['missed']
		else:
			Blinded_Rate = 0
		Blinded_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['missed'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])+' direct hit(s)">'+str(Blinded_Rate)+'</span>'
		
		output_string += str(Blinded_Rate_TT)+" | "
		
		#Calculate Interupt_Rate
		if squad_offensive[squadDps_prof_name]['stats']['interrupts']:
			Interupt_Rate = squad_offensive[squadDps_prof_name]['stats']['interrupts']
		else:
			Interupt_Rate = 0		
		Interupt_Rate_TT = '<span data-tooltip="Interupted enemy players '+str(Interupt_Rate)+' time(s)">'+str(Interupt_Rate)+'</span>'
		
		output_string += str(Interupt_Rate_TT)+" | "
		
		#Calculate Invulnerable_Rate
		if squad_offensive[squadDps_prof_name]['stats']['invulned']:
			Invulnerable_Rate = squad_offensive[squadDps_prof_name]['stats']['invulned']
		else:
			Invulnerable_Rate = 0
		Invulnerable_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['invulned'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])+' hit(s)">'+str(Invulnerable_Rate)+'</span>'
		
		output_string += str(Invulnerable_Rate_TT)+" | "
		
		#Calculate Evaded_Rate
		if squad_offensive[squadDps_prof_name]['stats']['evaded']:
			Evaded_Rate = squad_offensive[squadDps_prof_name]['stats']['evaded']
		else:
			Evaded_Rate = 0
		Evaded_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['evaded'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' direct hit(s)">'+str(Evaded_Rate)+'</span>'
		
		output_string += str(Evaded_Rate_TT)+" | "
		
		#Calculate Blocked_Rate
		if squad_offensive[squadDps_prof_name]['stats']['blocked']:
			Blocked_Rate = squad_offensive[squadDps_prof_name]['stats']['blocked']
		else:
			Blocked_Rate = 0		
		Blocked_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['blocked'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' direct hit(s)">'+str(Blocked_Rate)+'</span>'
		
		output_string += str(Blocked_Rate_TT)+" |"
		
		myprint(output, output_string)

	write_squad_offensive_xls(squad_offensive, args.xls_output_filename)
	myprint(output, "</$reveal>\n")
	#end Offensive Stat Table insert

	for stat in config.stats_to_compute:
		if stat == 'dist':
			write_stats_xls(players, top_percentage_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_percentage_stat_players[stat], stat, myDate, args.input_directory, config)
		elif stat == 'dmg_taken':
			write_stats_xls(players, top_average_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_average_stat_players[stat], stat, myDate, args.input_directory, config)
		elif stat == 'heal' and found_healing:
			write_stats_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_total_stat_players[stat], stat, myDate, args.input_directory, config)
		elif stat == 'barrier' and found_barrier:
			write_stats_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_total_stat_players[stat], stat, myDate, args.input_directory, config)
		elif stat == 'deaths':
			write_stats_xls(players, top_consistent_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_consistent_stat_players[stat], stat, myDate, args.input_directory, config)
		else:
			write_stats_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_total_stat_players[stat], stat, myDate, args.input_directory, config)
			if stat == 'rips' or stat == 'cleanses' or stat == 'stability':
				supportCount = write_support_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename, supportCount)