import requests
import json
import datetime
import win32com.client
from pathlib import Path


def get_weather_data(city, api_key):
    url = f"https://api.weatherapi.com/v1/current.json?key={api_key}&q={city}"

    try:
        req = requests.get(url)
        req.raise_for_status()  # Raise an exception for HTTP errors

        data = req.json()
        temp = data["current"]["temp_c"]
        print(f"The current temperature of {city} is {temp} degrees Celsius.")
        
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.Speak(f"The current temperature of {city} is {temp} degrees Celsius.")

        return data
    except requests.exceptions.RequestException as e:
        print(f"Error fetching weather data: {e}")
        return None


def save_weather_data(data):
    current_time = datetime.datetime.now()
    filename = Path(f"weather_data_{current_time.strftime('%Y-%m-%d_%H-%M-%S')}.json")

    try:
        with open(filename, 'w') as outfile:
            json.dump(data, outfile, indent=4)
        print(f"Weather data saved successfully to {filename}.")
    except IOError as e:
        print(f"Error saving weather data: {e}")


if __name__ == "__main__":
    api_key = "9b3c742e6b5c436ea9a124126232509"  # It's recommended to use environment variables or a config file
    city = input("Enter the city name: ")
    weather_data = get_weather_data(city, api_key)

    if weather_data:
        save_weather_data(weather_data)
    else:
        print("Failed to retrieve weather data.")
