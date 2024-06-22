TOTAL_MANGA = 19062

class Anime:
    def __init__(self, anime_data) -> None:
        """
        anime_data - data of the anime from MAL API

        anime_id - ID of the anime in MyAnimeList
        title - title of the anime
        favourites - number of users who have favorited the anime
        p2w - number of users who have added the anime to their Plan to Watch list
        watching - number of users who are currently watching the anime
        dropped - number of users who have dropped the anime
        completed - number of users who have completed the anime
        rating - rating of the anime
        """
        self.id = anime_data['id']
        self.title = anime_data['title']
        self.favourites =int(anime_data['num_favorites'])
        self.p2w = int(anime_data['statistics']['status']['plan_to_watch'])
        self.watching = int(anime_data['statistics']['status']['watching'])
        self.completed = int(anime_data['statistics']['status']['completed'])
        self.dropped = int(anime_data['statistics']['status']['dropped'])
        self.rating = anime_data.get("mean") or -1
        self.source = anime_data.get('source')

    def favourites_per_100_p2w(self) -> float:
        return self.favourites / self.p2w * 100

    def get_mal_link(self) -> str:
        return f"https://myanimelist.net/anime/{self.id}"

    def get_drop_rate(self) -> float:
        total = self.dropped + self.completed + self.watching
        if total == 0:
            return 0.0
        return self.dropped / total * 100

    def __repr__(self) -> str:
        return f"{self.title}"

class AdaptedAnime(Anime):

    def __init__(self, anime_data, manga_data) -> None:
        """
        manga_data - data of the manga from MAL API

        manga_id - ID of the manga in MyAnimeList
        manga_favourite - number of users who have favorited the manga
        manga_score - mean score of the manga
        manga_num_list_users - number of users who have added the manga to their list
        manga_rank - rank of the manga
        """
        super().__init__(anime_data)
        self.manga_id = manga_data['id']
        self.manga_favourite = manga_data['num_favorites']
        self.manga_score = manga_data.get('mean') or 0
        self.manga_num_list_users = manga_data['num_list_users']
        self.manga_rank = manga_data['rank']
        self.manga_type = manga_data['media_type']

    def get_manga_percentile(self) -> float:
        if self.manga_score <= 0:
            return -1
        return (1 - self.manga_rank / TOTAL_MANGA) * 100

    def get_manga_link(self) -> str:
        return f"https://myanimelist.net/manga/{self.manga_id}"

    def get_favourite_to_num_list_users(self) -> float:
        return self.manga_favourite / self.manga_num_list_users * 100


class SequelAnime(Anime):
    def __init__(self, anime_data, prequel_data, sequel_type: str, season: int) -> None:
        """
        prequel_data - data of the prequel from MAL API
        sequel_type - type of sequel (part 2, season 2, spin-off, etc.)
        season - season of the sequel (1, 2, 2.5, etc.)
        """
        super().__init__(anime_data)
        self.prequel = Anime(prequel_data)
        self.sequel_type = sequel_type
        self.season = season
