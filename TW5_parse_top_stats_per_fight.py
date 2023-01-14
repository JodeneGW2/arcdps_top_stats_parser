#!/usr/bin/env python3

from os import listdir
import argparse
import os.path
import sys
import importlib
import json
import xlwt
import datetime
import gzip

from TW5_parse_top_stats_tools import fill_config, reset_globals, get_stats_from_fight_json, get_stat_from_player_json, get_buff_ids_from_json, get_combat_time_breakpoints, sum_breakpoints

if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='This reads a set of arcdps reports in xml format and generates top stats.')
	parser.add_argument('input_directory', help='Directory containing .xml or .json files from arcdps reports')
	parser.add_argument('-c', '--config_file', dest="config_file", help="Config file with all the settings", default="TW5_parser_config_detailed")
	parser.add_argument('-a', '--anonymized', dest="anonymize", help="Create an anonymized version of the top stats. All account and character names will be replaced.", default=False, action='store_true')
	args = parser.parse_args()

	if not os.path.isdir(args.input_directory):
		print("Directory ",args.input_directory," is not a directory or does not exist!")
		sys.exit()

	log = open(args.input_directory+"/log_detailed.txt", "w")

	parser_config = importlib.import_module("parser_configs."+args.config_file , package=None) 
	config = fill_config(parser_config)
	
	# create xls file if it doesn't exist
	book = xlwt.Workbook(encoding="utf-8")
	sheet1 = book.add_sheet("Player Stats")

	sheet1.write(0, 0, "Account")
	sheet1.write(0, 1, "Name")
	sheet1.write(0, 2, "Profession")
	sheet1.write(0, 3, "Role")
	sheet1.write(0, 4, "Rally Num")
	sheet1.write(0, 5, "Fight Num")
	sheet1.write(0, 6, "Date")
	sheet1.write(0, 7, "Start Time")
	sheet1.write(0, 8, "End Time")
	sheet1.write(0, 9, "Num Allies")
	sheet1.write(0, 10, "Num Enemies")
	sheet1.write(0, 11, "Group")
	sheet1.write(0, 12, "Duration")
	sheet1.write(0, 13, "Combat time")
	sheet1.write(0, 14, "Damage")
	sheet1.write(0, 15, "Power Damage")
	sheet1.write(0, 16, "Condi Damage")
	sheet1.write(0, 17, "Crit Perc")
	sheet1.write(0, 18, "Flanking Perc")
	sheet1.write(0, 19, "Glancing Perc")
	sheet1.write(0, 20, "Blind Num")
	sheet1.write(0, 21, "Interrupt Num")
	sheet1.write(0, 22, "Invulnerable Num")
	sheet1.write(0, 23, "Evaded Num")
	sheet1.write(0, 24, "Blocked Num")
	sheet1.write(0, 25, "Coordination Damage")
	sheet1.write(0, 26, "Carrion Damage")
	for i in range(1, 21):
		sheet1.write(0, 26 + i, 'Chunk Damage (' + str(i) + ')')

	stats_to_compute = ['downs', 'kills', 'res', 'deaths', 'dmg_taken', 'barrierDamage', 'dist',  'swaps', 'rips', 'cleanses', 'barrier']
	for i,stat in enumerate(stats_to_compute):
		sheet1.write(0, 47+i, config.stat_names[stat])
		
	sheet1.write(0, 58, 'Total Healing')
	sheet1.write(0, 59, 'Power Healing')
	sheet1.write(0, 60, 'Conversion Healing')
	sheet1.write(0, 61, 'Hybrid Healing')
	sheet1.write(0, 62, 'Group Total Healing')
	sheet1.write(0, 63, 'Group Power Healing')
	sheet1.write(0, 64, 'Group Conversion Healing')
	sheet1.write(0, 65, 'Group Hybrid Healing')

	uptime_Order = ['stability',  'protection',  'aegis',  'might',  'fury',  'resistance',  'resolution',  'quickness',  'swiftness',  'alacrity',  'vigor',  'regeneration']
	for i,stat in enumerate(uptime_Order):
		sheet1.write(0, 66+i, stat.capitalize()+" Uptime")

	for i,stat in enumerate(uptime_Order):
		sheet1.write(0, 78+i, stat.capitalize()+" Gen")
			
	for i in range(0, 26):
		sheet1.write(0, 90 + i, 'Might (' + str(i) + ')')


	# iterating over all fights in directory
	files = listdir(args.input_directory)
	sorted_files = sorted(files)
	rally_num = 1
	fight_num = 1
	last_fight_end_time = None
	row = 1
	for filename in sorted_files:
		file_start, file_extension = os.path.splitext(filename)
		if file_extension not in ['.json', '.gz'] or "top_stats" in file_start:
			continue

		print("parsing "+filename)
		file_path = "".join((args.input_directory,"/",filename))

		if file_extension == '.gz':
			with gzip.open(file_path, mode="r") as f:
				json_data = json.loads(f.read().decode('utf-8'))
		else:
			json_datafile = open(file_path, encoding='utf-8')
			json_data = json.load(json_datafile)

		reset_globals()
		config = fill_config(parser_config)
		get_buff_ids_from_json(json_data, config)
		fight, players_running_healing_addon, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableIn, auras_TableOut, Death_OnTag, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats = get_stats_from_fight_json(json_data, config, log)
		
		if fight.skipped:
			continue


		if last_fight_end_time:
			after_last_fight = datetime.datetime.fromisoformat(last_fight_end_time) + datetime.timedelta(hours=2)
			if after_last_fight <  datetime.datetime.fromisoformat(fight.start_time):
				print("Start of a new rally at ", fight.start_time)
				rally_num += 1
				fight_num = 1

		last_fight_end_time = fight.end_time

		for squadDps_prof_name in DPSStats:
			player = [p for p in json_data["players"] if p['account'] == DPSStats[squadDps_prof_name]['account']][0]
			player_prof_name = "{{"+player['profession']+"}} "+player['name']

			fight_duration = json_data["durationMS"] / 1000
			combat_time = sum_breakpoints(get_combat_time_breakpoints(player)) / 1000

			# Calculate healing values
			if player['name'] in players_running_healing_addon and 'extHealingStats' in player:
				total_healing = 0
				total_healing_group = 0
				power_healing = 0
				power_healing_group = 0
				conversion_healing = 0
				conversion_healing_group = 0
				hybrid_healing = 0
				hybrid_healing_group = 0

				allied_healing_1s = player['extHealingStats']['alliedHealing1S']
				allied_power_healing_1s = player['extHealingStats']['alliedHealingPowerHealing1S']
				allied_conversion_healing_1s = player['extHealingStats']['alliedConversionHealingHealing1S']
				allied_hybrid_healing_1s = player['extHealingStats']['alliedHybridHealing1S']
				for index in range(len(json_data['players'])):
					is_same_group = player['group'] == json_data['players'][index]['group']

					player_healing = allied_healing_1s[index][0][-1]
					player_power_healing = allied_power_healing_1s[index][0][-1]
					player_conversion_healing = allied_conversion_healing_1s[index][0][-1]
					player_hybrid_healing = allied_hybrid_healing_1s[index][0][-1]


					total_healing += player_healing
					power_healing += player_power_healing
					conversion_healing += player_conversion_healing
					hybrid_healing += player_hybrid_healing
					if is_same_group:
						total_healing_group += player_healing
						power_healing_group += player_power_healing
						conversion_healing_group += player_conversion_healing
						hybrid_healing_group += player_hybrid_healing
			else:
				# When no healing stats data, set to -1
				total_healing = -1
				total_healing_group = -1
				power_healing = -1
				power_healing_group = -1
				conversion_healing = -1
				conversion_healing_group = -1
				hybrid_healing = -1
				hybrid_healing_group = -1

			sheet1.write(row, 0, DPSStats[squadDps_prof_name]['account'])
			sheet1.write(row, 1, DPSStats[squadDps_prof_name]['name'])
			sheet1.write(row, 2, DPSStats[squadDps_prof_name]['profession'])
			sheet1.write(row, 3, DPSStats[squadDps_prof_name]['role'])
			sheet1.write(row, 4, rally_num)
			sheet1.write(row, 5, fight_num)
			sheet1.write(row, 6, fight.start_time.split()[0])
			sheet1.write(row, 7, fight.start_time.split()[1])
			sheet1.write(row, 8, fight.end_time.split()[1])
			sheet1.write(row, 9, fight.allies)
			sheet1.write(row, 10, fight.enemies)
			sheet1.write(row, 11, int(player['group']))
			sheet1.write(row, 12, fight_duration)
			sheet1.write(row, 13, combat_time)
			sheet1.write(row, 14, get_stat_from_player_json(player, players_running_healing_addon, 'dmg', config))
			sheet1.write(row, 15, get_stat_from_player_json(player, players_running_healing_addon, 'Pdmg', config))
			sheet1.write(row, 16, get_stat_from_player_json(player, players_running_healing_addon, 'Cdmg', config))

			if squad_offensive[player_prof_name]['stats']['critableDirectDamageCount'] > 0:
				sheet1.write(row, 17, squad_offensive[player_prof_name]['stats']['criticalRate'] / squad_offensive[player_prof_name]['stats']['critableDirectDamageCount'])
			else:
				sheet1.write(row, 17, 0)

			if squad_offensive[player_prof_name]['stats']['connectedDirectDamageCount'] > 0:
				sheet1.write(row, 18, squad_offensive[player_prof_name]['stats']['flankingRate'] / squad_offensive[player_prof_name]['stats']['connectedDirectDamageCount'])
				sheet1.write(row, 19, squad_offensive[player_prof_name]['stats']['glanceRate'] / squad_offensive[player_prof_name]['stats']['connectedDirectDamageCount'])
			else:
				sheet1.write(row, 18, 0)
				sheet1.write(row, 19, 0)

			sheet1.write(row, 20, squad_offensive[player_prof_name]['stats']['missed'])
			sheet1.write(row, 21, squad_offensive[player_prof_name]['stats']['interrupts'])
			sheet1.write(row, 22, squad_offensive[player_prof_name]['stats']['invulned'])
			sheet1.write(row, 23, squad_offensive[player_prof_name]['stats']['evaded'])
			sheet1.write(row, 24, squad_offensive[player_prof_name]['stats']['blocked'])
			sheet1.write(row, 25, DPSStats[squadDps_prof_name]['Coordination_Damage'])
			sheet1.write(row, 26, DPSStats[squadDps_prof_name]['Carrion_Damage'])

			for i in range(1, 21):
				sheet1.write(row, 26 + i, DPSStats[squadDps_prof_name]['Chunk_Damage'][i])

			for i,stat in enumerate(stats_to_compute):
				sheet1.write(row, 47+i, get_stat_from_player_json(player, players_running_healing_addon, stat, config))
		
			sheet1.write(row, 58, total_healing)
			sheet1.write(row, 59, power_healing)
			sheet1.write(row, 60, conversion_healing)
			sheet1.write(row, 61, hybrid_healing)
			sheet1.write(row, 62, total_healing_group)
			sheet1.write(row, 63, power_healing_group)
			sheet1.write(row, 64, conversion_healing_group)
			sheet1.write(row, 65, hybrid_healing_group)

			for i,stat in enumerate(uptime_Order):
				if stat in uptime_Table[player_prof_name]:
					buff_Time = uptime_Table[player_prof_name][stat]
					sheet1.write(row, 66+i, buff_Time)
				else:
					sheet1.write(row, 66+i, 0.00)

			for i,stat in enumerate(uptime_Order):
				value = get_stat_from_player_json(player, players_running_healing_addon, stat, config)
				if stat in config.buffs_stacking_duration:
					sheet1.write(row, 78+i, value/100.*fight_duration)
				elif stat in config.buffs_stacking_intensity:
					sheet1.write(row, 78+i, value*fight_duration)
			
			for i in range(0, 26):
				sheet1.write(row, 90 + i, stacking_uptime_Table[squadDps_prof_name]['might'][i] / 1000.0)

			row += 1
		
		fight_num += 1

	book.save(args.input_directory+"/TW5_top_stats_per_fight.xls")