#Here is the code to provide temprature updates of a city by using openweather API 
import requests
from pprint import pprint

def weather_data(query):
    res=requests.get('https://api.openweathermap.org/data/2.5/weather?'+query+'&appid=c8b92668120f071c0ee81409f2f1ad54'); #openweather site is called here. 
    return res.json();
values = []
def print_weather(result,city):
    temp = "{}'s temperature: {}°C ".format(city,result['main']['temp'])
    values.append(temp)
    wind_speed = "Wind speed: {} m/s".format(result['wind']['speed'])
    values.append(wind_speed)
    wea_forecast = "Description: {}".format(result['weather'][0]['description'])
    values.append(wea_forecast)
    conditions = "Weather: {}".format(result['weather'][0]['main'])
    values.append(conditions)
    print(temp)
    print(wind_speed)
    print(wea_forecast)
    print(conditions)

city=input('Enter the city:')# Enter the name of the city
print()
try:
    query='q='+city;
    w_data=weather_data(query);
    print_weather(w_data, city)
    print()
except:
    print('City name not found...')
    
val_real = [city]
for i in range(len(values)):
    for j in range(len(values[i])):
        if(values[i][j] == ":"):
            flag = j
    x = values[i][flag+2:len(values[i])]
    val_real.append(x)

Cols = ["A","B","C","D","E"]
initials = ["City","Temperature","Wind Speed","Forecast","Conditions"]

from openpyxl import Workbook#imported pyxl library to update temprture reports to the excel sheet

workbook = Workbook()
file = workbook.active

for i in range(len(Cols)):
    file[Cols[i]+str(1)] = initials[i]
    file[Cols[i]+str(2)] = val_real[i]
workbook.save(filename="WeatherAPI.xlsx")
