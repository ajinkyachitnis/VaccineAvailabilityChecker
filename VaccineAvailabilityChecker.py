#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon May  11 08:10:21 2021
@author: Ajinkya Chintis
"""

import requests
import json
import time
import datetime
from win32com.client import Dispatch
from twilio.rest import Client

speak = Dispatch("SAPI.SpVoice").Speak

pincode = '<ENTER_PINCODE HERE>' #Enter Your Area Pincode

while True:
    now = datetime.datetime.now()
    print('=========================Slots Status=========================')
    print ("Current date and time : ", now.strftime("%d-%m-%Y %H:%M:%S"))
    date = now.strftime("%d-%m-%Y")
    response =requests.get('https://cdn-api.co-vin.in/api/v2/appointment/sessions/public/calendarByPin',
                           params={'pincode':pincode,'date':date},
                           headers={'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36'}) #requet for cowin availability API
    #save result in json format
    if str(response) == '<Response [200]>':        
        json_response = response.json()
        all_centers = json_response['centers']
        count = 0
        #logic to parse json and obtain desired results
        for center in all_centers:
            center_location = center['block_name']
            center_name=center['name']
            all_sessions=center['sessions']
            for session in all_sessions:
                available_slot=session['available_capacity']
                age_limit=session['min_age_limit']
                if age_limit == 18:
                    if available_slot != 0:
                        count = count + 1
                        print('***********************Slots Available***********************')
                        slot_details = 'Age Limit is :' + str(age_limit) +'\n' + 'Available Slots are :' + str(available_slot) + '\n at'
                        print(slot_details)
                        print(center_name)
                        for i in range(3):
                            slot_info = str(available_slot) + ' Vaccine Slots Available at' + center_name
                            speak(slot_info)
                        #twilio supported code to send text message, can be committed out if not required    
                        if count <= 1 :
                            client = Client('<TWILIO_ACCOUNT_SID>', '<TWILIO_AUTH_TOKEN>')
                            message = client.messages \
                                .create(body=slot_info,from_='<TWILIO_PHONE_NUMBER>',to='+<YOUR_PHONE_NUMBER')   
                    else :
                        print(center_name,': No Slot Available Yet \n')
    else:
        print('Request blocked.We cannot connect to the server for this app or website at this time. There might be too much traffic.')
        time.sleep(60) #if request fails to get 200 OK response wait for 60 seconds
    time.sleep(10) #delay given as sometimes rapidly calling API results to higher traffic and leads to crash.     
           
