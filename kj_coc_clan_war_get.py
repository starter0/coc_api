#! /usr/bin/env python3


import asyncio
import jwt
import requests, json
import datetime, time
import logging
import kj_coc_lib
import kj_war_to_excel as wte


request_interval = 300
my_clan_tag='#982JCPJU'
my_key_name = "Created with coc.py Client" 
my_key = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiIsImtpZCI6IjI4YTMxOGY3LTAwMDAtYTFlYi03ZmExLTJjNzQzM2M2Y2NhNSJ9.eyJpc3MiOiJzdXBlcmNlbGwiLCJhdWQiOiJzdXBlcmNlbGw6Z2FtZWFwaSIsImp0aSI6IjU2Yjc4NWQzLWRiNDYtNGQwZS04N2VhLWY3NGY1Y2RmN2MxMiIsImlhdCI6MTU4MzgwMzM1MCwic3ViIjoiZGV2ZWxvcGVyLzY5ZDVjYmRiLTg5NTQtMDc4Zi1hNzYwLTY2ZmY4MDFiMTc3MyIsInNjb3BlcyI6WyJjbGFzaCJdLCJsaW1pdHMiOlt7InRpZXIiOiJkZXZlbG9wZXIvc2lsdmVyIiwidHlwZSI6InRocm90dGxpbmcifSx7ImNpZHJzIjpbIjIyMC4yMzAuMTg0LjEzNSJdLCJ0eXBlIjoiY2xpZW50In1dfQ.N2ntG4IN3nuEOO5_rFtA0SRlo5nA2XDMdc3SxwyfcCfXBbB1T6yP29MHAsJri39zLylO14E7A_SvtHFcpDPYCA"

my_key2= "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiIsImtpZCI6IjI4YTMxOGY3LTAwMDAtYTFlYi03ZmExLTJjNzQzM2M2Y2NhNSJ9.eyJpc3MiOiJzdXBlcmNlbGwiLCJhdWQiOiJzdXBlcmNlbGw6Z2FtZWFwaSIsImp0aSI6ImYzNjM2NDFiLWRkMTEtNDkwOS1hOGRmLTlhMzAyMmUyYmE2YSIsImlhdCI6MTU4NDQwNDg3OSwic3ViIjoiZGV2ZWxvcGVyLzY5ZDVjYmRiLTg5NTQtMDc4Zi1hNzYwLTY2ZmY4MDFiMTc3MyIsInNjb3BlcyI6WyJjbGFzaCJdLCJsaW1pdHMiOlt7InRpZXIiOiJkZXZlbG9wZXIvc2lsdmVyIiwidHlwZSI6InRocm90dGxpbmcifSx7ImNpZHJzIjpbIjE4Mi4yMjkuODAuMjAiXSwidHlwZSI6ImNsaWVudCJ9XX0.ehu1yYEbd7NkbuEJ6NGHF1AzRXcY9tTTEvbkr8xKdMlzTMqJ_tkq11LnBUsQwAefMaG-cPz0DvWJMRI3p0j7Rg"


def kj_api_login() :
	coc_id = '80dack@naver.com'
	coc_passwd = 'dgenius11'

	login_page = "https://developer.clashofclans.com/api/login"

	login_data = {"email": coc_id, "password": coc_passwd}
	headers = {"content-type": "application/json"}

	my_logger.debu('Login Try!!!') 
	response = requests.post(login_page, headers = headers, json = login_data)
	my_logger.debu('Login result is'+str(response.status_code))

	return response


pre_state = ''

while True:

	mylogger = kj_coc_lib.create_logger('kj_coc_log')

	kj_coc_lib.coc_login_kj()
	mylogger.debug('Did login!!')
	
	#auth = { 'Authorization': 'Bearer {}'.format(my_key2) }
	request_url = 'https://api.clashofclans.com/v1/clans/{clanTag}/currentwar'.format(clanTag = requests.utils.quote(my_clan_tag))

	s = requests.Session()

	res = s.get(request_url, headers={'Authorization': 'Bearer %s' %my_key2})

	now = datetime.datetime.now()
	current_time = now.strftime('%Y-%m-%d_%H_%M_%S')

	#print ('GET currenwar of clans response code is 'res.status_code)
	mylogger.debug('GET current war of clans  response code is '+str(res.status_code))

	if ( res.status_code == 200) :
		res_txt = json.loads(res.content.decode('utf-8'))
		if (res_txt['state'] == 'warEnded') :
			if (pre_state == 'warEnded') :
				mylogger.debug('Not save file cause it saved already')
				time.sleep(60*10)
				continue
				
			else: 
				with open(current_time+'_war_result.json','w', encoding='utf-8') as make_file: 
					json.dump(res_txt, make_file, indent="\t")
				mylogger.debug('saved file cause it is warEnded')
				mylogger.debug('Try to make excel file')
				pre_state = 'warEnded'
				wte.war_to_excel(res_txt)

		else :
			pre_state = res_txt['state']
			mylogger.debug('Not save file cause it is not warended status. current state is '+pre_state)
			#print ('Not save file cause it is not warended status')	
			#log_file.write(datetime.datetime.now() + '\t Not save file')

	else : 
		mylogger.debug('Failed to access clan_war \t response code is '+str(res.status_code))
		print ('Failed to access clan_war \t response code is '+str(res.status_code))
		if ( res.status_code >= 500) :
			mylogger.debug('Sleep {} seconds'.format(str(request_interval*5)))
			time.sleep(request_interval*5)
			mylogger.debug('Try to relogin')
			kj_coc_lib.coc_login_kj()
			


	if ( pre_state == 'inWar') :
		mylogger.debug('Try to sleep for %s seconds cause state is inWar'%request_interval)
		time.sleep(request_interval)
		
	else :
	
		mylogger.debug('Try to sleep for {interval} seconds cause state is {state}'.format(interval = request_interval*12*60, state = pre_state))
		time.sleep(request_interval*12*6)


