import requests
from dotenv import load_dotenv
import os
from anime_data_model import Anime

load_dotenv()
CLIENT_ID = os.getenv("CLIENT_ID")

def fetch_anime_data(anime_id):
    url = f'https://api.myanimelist.net/v2/anime/{anime_id}?fields=mean,num_favorites,statistics'
    response = requests.get(url, headers={'X-MAL-CLIENT-ID': CLIENT_ID})
    response.raise_for_status()
    anime_data = response.json()
    print(anime_data)
    response.close()
    return anime_data

def fetch_anime_season_data(year, season):
    url = f'https://api.myanimelist.net/v2/anime/season/{year}/{season}?limit=4'
    response = requests.get(url, headers={'X-MAL-CLIENT-ID': CLIENT_ID})
    response.raise_for_status()
    anime_season_data = response.json()
    for anime in anime_season_data['data']:
        anime_id = anime['node']['id']
        fetch_anime_data(anime_id)
    response.close()

def main():
    fetch_anime_season_data(2023, 'winter')

if __name__ == "__main__":
    main()
