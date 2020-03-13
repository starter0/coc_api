#! /usr/bin/env python3

import os, sys, json
import subprocess
import datetime, time

import openpyxl
from openpyxl.styles import Font, Alignment


def member_arrange(member_obj) :

	arranged_member_list = {}

	for member in member_obj['members'] :
		arranged_member_list.update({member['tag'] : member['mapPosition']} )

	return (arranged_member_list)		

def xlsx_idle(path, filename_xls, op_clan) :
	
    if os.path.isfile(path+ filename_xls) :
    #if os.path.isfile(path+'mem_info.xlsx') :
        wb = openpyxl.load_workbook(path+filename_xls)
        #wb = openpyxl.load_workbook(path+'mem_info.xlsx')
    else :
        wb = openpyxl.Workbook()
    
    sh = wb.create_sheet(str(datetime.datetime.now().date()), 0)
    sh['B1'] = 'VS ' + op_clan
    sh['B2'] = 'Member Name' 
    sh['C2'] = 'Position'
    sh['D2'] = '1th attack position'
    sh['E2'] = 'stars'
    sh['F2'] = 'Destruction Percentage'
    sh['G2'] = '2st attack position'
    sh['H2'] = 'stars'
    sh['I2'] = 'DestructionPercentage'
    sh['J2'] = 'Not used attack count'
    sh['K2'] = 'Total stars'
    sh['L2'] = 'Town hall level'
    sh.merge_cells('M1:O1')
    sh['M1'] = 'Best opponent attack'
    sh['M2'] = 'Attacker position'
    sh['N2'] = 'stars'
    sh['O2'] = 'DestructionPercentage' 
    sh['P2'] = 'opponentAttacks'
    #sh['L3'] = 'Baba king'
    #sh['M3'] = 'Archer queen'
    #sh['N3'] = 'Warden'
    #sh['O3'] = 'Royal champion'
    
    
    # style
    #sh.freeze_panes = 'A4'
    sh['B1'].font = Font(size = 20, bold = True)

    wb.save(path+filename_xls)
"""
    rows = sh.max_row
    cols = sh.max_column
    for r in range(1, rows) :
        for c in range(0,cols) :
            sh.cell(row = r+1, column = c+1).alignment = Alignment(horizontal='center', vertical='center')
            sh.cell(row = r+1, column = c+1).font = Font(bold = True)

"""
	

def append_xlsx(path, obj, filename_xls) :
    wb = openpyxl.load_workbook(path + filename_xls)
    sh = wb.active
    sh.append(obj)
    wb.save(path + filename_xls)

def sort_xlsx(path,filename_xls) : # not worked yet
	wb = openpyxl.load_workbook(path + filename_xls)
	sh = wb.active
	sh.auto_filter.add_sort_condition('B2:L27',False)
	wb.save(path + filename_xls)



if __name__ == '__main__' :

	if ( len(sys.argv) < 2) :
		print ('Usage is')
		exit (1)	

	with open(sys.argv[1]) as json_file:
		json_data = json.load(json_file)

	opponent_member = member_arrange(json_data['opponent'])
	opponent_clan_name = json_data['opponent']['name']

	title = 'Clan_war.xlsx'	

	xlsx_idle('./',title,opponent_clan_name)
	append_obj = ['']

	for member in json_data['clan']['members'] :
		append_obj.append(member['name'])
		append_obj.append(member['mapPosition'])

		
		if ( 'attacks' in member ) :
			for attack in member['attacks'] :
				#find opponent postion
				append_obj.append(opponent_member[(attack['defenderTag'])])
				append_obj.append(attack['stars'])
				append_obj.append(attack['destructionPercentage'])

	
		################ To do 
		# caculate miised attack count
		if ( len (append_obj) < 4 ) : # If not attacked
			remain_count = 2
			total_stars = 0
			append_obj.extend(('','','','','',''))

		elif ( len (append_obj) < 7 ) : # If attacked 1 time
			remain_count = 1
			total_stars = append_obj[4]
			append_obj.extend(('','',''))
		else :
			remain_count = 0
			total_stars = int(append_obj[4]) + int(append_obj[7])

		append_obj.append(remain_count)

		# Total stars

		append_obj.append(total_stars)
		
		# town hall
		append_obj.append(member['townhallLevel'])
		
		# Heroes level baba king
			

		################### To do end
	
		### Best opponent attack info 
		if ( 'opponentAttacks' in member ) :
			#opponent_position = opponent_member[member['bestOpponentAttack']['attackerTag']]
			append_obj.append(opponent_member[member['bestOpponentAttack']['attackerTag']])
			append_obj.append(member['bestOpponentAttack']['stars'])
			append_obj.append(member['bestOpponentAttack']['destructionPercentage'])
			append_obj.append(member['opponentAttacks'])

		else : 
			append_obj.extend('','','','')

		append_xlsx('./',append_obj,title)

		append_obj = ['']


