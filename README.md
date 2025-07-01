# Fantasy Anime League (FAL)
Retrieving data for Fantasy Anime League in order to predict how well Anime will do on the upcoming season.

## Pre-season Data Scraper usage
copy .env_example to .env and add MAL api token as CLIENT_ID.

run data_scraper/season_scraper.py to generate a csv containing shows of the season

remove ids of banned shows. Add any shows id the scraper missed (~5% are missed. Usually small shows or "gray" rated shows)

run data_scraper/season_scraper.py with variable for_fal, remove_ids and add_ids updated to generate a csv containing shows of the FAL season

for all non-sequel shows, manually find whether the show is an adaptation and find the id of the manga

populate the ids of sequels adaptations and originals

run data_scraper/scraper_based_on_anime_type.py

A spreadsheet with 3 sheets named original, adapatation and sequel should be made. Example of what adaptation sheet should look like:

![image](https://github.com/user-attachments/assets/744dee21-362e-43c7-b63c-28b24c652543)
