# Fantasy Anime League (FAL)
Retrieving data for Fantasy Anime League in order to predict how well Anime will do on the upcoming season.

note: this repository is not well organized nor well documented as
it was made quickly before FAL start
and during start of a school term.

To-do: organize the repository and increase automation to minimize work between FAL seasons to acquire data.
should also get air date of last prequel

<<<<<<< HEAD
# Pre-season Data Scraper usage
copy .env_example to .env and add MAL api token as CLIENT_ID.
run data_scraper/season_scraper.py to generate a csv containing shows of the season
remove ids of banned shows. Add any shows id the scraper missed (~5% are missed. Usually small shows or "gray" rated shows)
run data_scraper/season_scraper.py with variable for_fal, remove_ids and add_ids updated to generate a csv containing shows of the FAL season
for all non-sequel shows, manually find whether the show is an adaptation and find the id of the manga
populate the ids of sequels adaptations and originals
run data_scraper/scraper_based_on_anime_type.py

=======
This will hopefully be updated next FAL season
# Usage
Copy .env_example to .env and add MAL api token as CLIENT_ID.

run season_scraper.py to generate a csv contataining all the shows

???

(clean data)

???

run scraper_based_on_anime_type to generate a csv obtaining data based on shows whether they are an adaptation, orginal or sequel
>>>>>>> d2b7939131d2f5159ea9e26f44bded03b9246982

? run scraper.py or main_season.py to get general data on shows.
