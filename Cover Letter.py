# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
api_key = "test"

import googlemaps
from googlemaps import GoogleMaps
gmaps = GoogleMaps(api_key)
address = 'Constitution Ave NW & 10th St NW, Washington, DC'
lat, lng = gmaps.address_to_latlng(address)
print(lat, lng)