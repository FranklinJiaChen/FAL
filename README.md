# FAL
Retrieving data for Fantasy Anime League.

note: this repository is not well organized nor well documented as
it was made quickly before FAL start
and during start of a school term.

To-do: organize the repository and increase automation to minimize work between FAL seasons to acquire data.
should also get air date of last prequel

# Pre-season Data Scraper usage
Copy .env_example to .env and add MAL api token as CLIENT_ID.
run data_scraper/season_scraper.py to generate a csv containing shows of the season
Remove ids of banned shows. Add any shows id the scraper missed (~5% are missed. Usually small shows or "gray" rated shows)
run data_scraper/season_scraper.py with variable for_fal, remove_ids and add_ids updated to generate a csv containing shows of the FAL season

???
(clean data)
???
run scraper_based_on_anime_type to generate a csv obtaining data based on shows whether they are an adaptation, orginal or sequel
? run scraper.py or main_season.py to get general data on shows.
