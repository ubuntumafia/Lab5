using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2019.Presentation;
using System;
using System.Collections.Generic;

public class Operations
{
    public string PathToXlsFile { get; }
    public Database Database { get; }

    public Operations(string pathToXlsFile)
    {
        PathToXlsFile = pathToXlsFile;
        Database = new Database(PathToXlsFile);
    }

    public void DisplayAlbums()
    {
        foreach (var album in Database.Albums)
        {
            Console.WriteLine(album);
        }
        Console.WriteLine();
    }

    public void DisplayArtists()
    {
        foreach (var artist in Database.Artists)
        {
            Console.WriteLine(artist);
        }
        Console.WriteLine();
    }

    public void DisplayTracks()
    {
        foreach (var track in Database.Tracks)
        {
            Console.WriteLine(track);
        }
        Console.WriteLine();
    }

    public void DisplayGenres()
    {
        foreach (var genre in Database.Genres)
        {
            Console.WriteLine(genre);
        }
        Console.WriteLine();
    }


    public void DeleteAlbumById(int id)
    {
        Album album = Database.Albums.Find(a => a.ID == id);
        if (album != null)
        {
            // Удаляем все треки, связанные с этим альбомом
            Database.Tracks.RemoveAll(t => t.AlbumID == album.ID);
            // Удаляем сам альбом
            Database.Albums.Remove(album);
            Console.WriteLine($"Альбом '{album.Name}' и все его треки удалены.");
            logger.LogAction($"Пользователь удалил альбом '{album.Name}'.");
        }
        else
        {
            Console.WriteLine($"Альбом с ID {id} не найден.");
            logger.LogAction("Пользователь пытался удалить альбом.");
        }
    }

    public void DeleteArtistById(int id)
    {
        Artist artist = Database.Artists.Find(a => a.ID == id);
        if (artist != null)
        {
            Database.Albums.RemoveAll(a => a.ArtistID == artist.ID);
            Database.Tracks.RemoveAll(t => Database.Albums.Any(a => a.ArtistID == artist.ID && a.ID == t.AlbumID));
            Database.Artists.Remove(artist);
            Console.WriteLine($"Артист '{artist.Name}', его альбомы и треки удалены.");
            logger.LogAction($"Пользователь удалил артиста '{artist.Name}'.");
        }
        else
        {
            Console.WriteLine($"Артист с ID {id} не найден.");
            logger.LogAction($"Пользователь пытался удалить артиста.");
        }
    }

    public void DeleteTrackById(int id)
    {
        Track track = Database.Tracks.Find(t => t.ID == id);
        if (track != null)
        {
            Database.Tracks.Remove(track);
            Console.WriteLine($"Трек '{track.Name}' удален.");
            logger.LogAction($"Пользователь удалил трек '{track.Name}'.");
        }
        else
        {
            Console.WriteLine($"Трек с ID {id} не найден.");
            logger.LogAction($"Пользователь пытался удалить артиста.");
        }
    }

    public void DeleteGenreById(int id)
    {
        Genre genre = Database.Genres.Find(g => g.ID == id);
        if (genre != null)
        {
            Database.Tracks.RemoveAll(t => t.GenreID == genre.ID);
            Database.Genres.Remove(genre);
            Console.WriteLine($"Жанр '{genre.Name}' и все его треки удалены.");
            logger.LogAction($"Пользователь удалил жанр '{genre.Name}'.");
        }
        else
        {
            Console.WriteLine($"Жанр с ID {id} не найден.");
            logger.LogAction($"Пользователь пытался удалить жанр.");
        }
    }

    
    public void UpdateAlbum()
    {
        int albumId = GetIntInput("Введите ID альбома для изменения:", 1, int.MaxValue);

        Album album = Database.Albums.Find(a => a.ID == albumId);
        if (album == null)
        {
            Console.WriteLine($"Альбом с ID {albumId} не найден.");
            logger.LogAction($"Пользователь пытался изменить альбом.");
            return;
        }

        string newName = GetStringInput($"Введите новое название альбома (текущее: {album.Name}) или 0 для пропуска:");
        if (newName != "0")
        {
            album.Name = newName;
            logger.LogAction($"Пользователь изменил название альбома {albumId} на '{album.Name}'.");
        }

        int newArtistId = GetIntInput($"Введите новый ID артиста (текущий: {album.ArtistID}) или 0 для пропуска:", 0, int.MaxValue);
        if (newArtistId != 0)
        {
            album.ArtistID = newArtistId;
            logger.LogAction($"Пользователь изменил ID артиста альбома {albumId} на '{album.ArtistID}'.");
        }

        Console.WriteLine("Альбом  обновлен.");
    }

    public void UpdateArtist()
    {
        int artistId = GetIntInput("Введите ID артиста для изменения:", 1, int.MaxValue);

        Artist artist = Database.Artists.Find(a => a.ID == artistId);
        if (artist == null)
        {
            Console.WriteLine($"Артист с ID {artistId} не найден.");
            logger.LogAction($"Пользователь пытался изменить артиста.");
            return;
        }

        string newName = GetStringInput($"Введите новое название артиста (текущее: {artist.Name}) или 0 для пропуска:");
        if (newName != "0")
        {
            artist.Name = newName;
            logger.LogAction($"Пользователь изменил название артиста {artistId} на '{artist.Name}'.");
        }

        Console.WriteLine("Артист обновлен.");
    }

    public void UpdateTrack()
    {
        int trackId = GetIntInput("Введите ID трека для изменения:", 1, int.MaxValue);

        Track track = Database.Tracks.Find(t => t.ID == trackId);
        if (track == null)
        {
            Console.WriteLine($"Трек с ID {trackId} не найден.");
            logger.LogAction($"Пользователь пытался изменить трек.");
            return;
        }

        string newName = GetStringInput($"Введите новое название трека (текущее: {track.Name}) или 0 для пропуска:");
        if (newName != "0")
        {
            track.Name = newName;
            logger.LogAction($"Пользователь изменил название трека {trackId} на '{track.Name}'.");
        }

        int newAlbumId = GetIntInput($"Введите новый ID альбома (текущий: {track.AlbumID}) или 0 для пропуска:", 0, int.MaxValue);
        if (newAlbumId != 0)
        {
            track.AlbumID = newAlbumId;
            logger.LogAction($"Пользователь изменил ID альбома трека {trackId} на '{track.AlbumID}'.");
        }

        int newGenreId = GetIntInput($"Введите новый ID жанра (текущий: {track.GenreID}) или 0 для пропуска:", 0, int.MaxValue);
        if (newGenreId != 0)
        {
            track.GenreID = newGenreId;
            logger.LogAction($"Пользователь изменил ID жанра трека {trackId} на '{track.GenreID}'.");
        }

        long newDuration = GetLongInput($"Введите новую длительность трека (текущая: {track.Duration}) или 0 для пропуска:");
        if (newDuration != 0)
        {
            track.Duration = newDuration;
            logger.LogAction($"Пользователь изменил продолжительность трека {trackId} на '{track.Duration}'.");
        }

        int newSize = GetIntInput($"Введите новый размер трека (текущий: {track.Size}) или 0 для пропуска:", 0, int.MaxValue);
        if (newSize != 0)
        {
            track.Size = newSize;
            logger.LogAction($"Пользователь изменил размер трека {trackId} на '{track.Size}'.");
        }

        decimal newCost = GetDecimalInput($"Введите новую стоимость трека (текущая: {track.Cost}) или 0 для пропуска:");
        if (newCost != 0)
        {
            track.Cost = newCost;
            logger.LogAction($"Пользователь изменил стоимость трека {trackId} на '{track.Cost}'.");
        }

        Console.WriteLine("Трек обновлен.");
    }

    public void UpdateGenre()
    {
        int genreId = GetIntInput("Введите ID жанра для изменения:", 1, int.MaxValue);

        Genre genre = Database.Genres.Find(g => g.ID == genreId);
        if (genre == null)
        {
            Console.WriteLine($"Жанр с ID {genreId} не найден.");
            logger.LogAction($"Пользователь пытался изменить жанр.");
            return;
        }

        string newName = GetStringInput($"Введите новое название жанра (текущее: {genre.Name}) или 0 для пропуска:");
        if (newName != "0")
        {
            genre.Name = newName;
            logger.LogAction($"Пользователь изменил название жанра {genreId} на '{genre.Name}'.");
        }

        Console.WriteLine($"Жанр обновлен.");
    }

    public void AddAlbum()
    {
        Console.WriteLine("Введите название нового альбома (оставьте пустым, чтобы не менять):");
        string newAlbumName = Console.ReadLine();

        Console.WriteLine("Введите ID нового артиста:");
        string newArtistId = Console.ReadLine();

        if (!string.IsNullOrEmpty(newAlbumName) || !string.IsNullOrEmpty(newArtistId))
        {
            int artistId = int.Parse(newArtistId);
            Database.Albums.Add(new Album(Database.Albums.Max(a => a.ID) + 1, newAlbumName, artistId));
            Console.WriteLine($"Альбом '{newAlbumName}' добавлен.");
            logger.LogAction($"Пользователь добавил альбом '{newAlbumName}'.");
        }
    }

    public void AddArtist()
    {
        Console.WriteLine("Введите название нового артиста:");
        string newArtistName = Console.ReadLine();

        if (!string.IsNullOrEmpty(newArtistName))
        {
            Database.Artists.Add(new Artist(Database.Artists.Max(a => a.ID) + 1, newArtistName));
            Console.WriteLine($"Артист '{newArtistName}' добавлен.");
            logger.LogAction($"Пользователь добавил исполнителя '{newArtistName}'.");
        }
    }

    public void AddTrack()
    {
        Console.WriteLine("Введите название нового трека:");
        string newTrackName = Console.ReadLine();

        Console.WriteLine("Введите ID альбома:");
        string newAlbumId = Console.ReadLine();

        Console.WriteLine("Введите ID жанра:");
        string newGenreId = Console.ReadLine();

        Console.WriteLine("Введите длительность трека:");
        string newDuration = Console.ReadLine();

        Console.WriteLine("Введите размер трека:");
        string newSize = Console.ReadLine();

        Console.WriteLine("Введите цену трека:");
        string newCost = Console.ReadLine();

        if (!string.IsNullOrEmpty(newTrackName) && !string.IsNullOrEmpty(newAlbumId) && !string.IsNullOrEmpty(newGenreId) && !string.IsNullOrEmpty(newDuration) && !string.IsNullOrEmpty(newSize) && !string.IsNullOrEmpty(newCost))
        {
            int albumId = int.Parse(newAlbumId);
            int genreId = int.Parse(newGenreId);
            long duration = long.Parse(newDuration);
            int size = int.Parse(newSize);
            decimal cost = decimal.Parse(newCost);

            Database.Tracks.Add(new Track(Database.Tracks.Max(t => t.ID) + 1, newTrackName, albumId, genreId, duration, size, cost));
            Console.WriteLine($"Трек '{newTrackName}' добавлен.");
            logger.LogAction($"Пользователь добавил исполнителя '{newTrackName}'.");
        }
    }

    public void AddGenre()
    {
        Console.WriteLine("Введите название нового жанра:");
        string newGenreName = Console.ReadLine();

        if (!string.IsNullOrEmpty(newGenreName))
        {
            Database.Genres.Add(new Genre(Database.Genres.Max(g => g.ID) + 1, newGenreName));
            Console.WriteLine($"Жанр '{newGenreName}' добавлен.");
            logger.LogAction($"Пользователь добавил жанр '{newGenreName}'.");
        }
    }

    public int GetIntInput(string prompt, int min, int max)
    {
        int input;
        while (true)
        {
            Console.Write(prompt);
            if (int.TryParse(Console.ReadLine(), out input) && input >= min && input <= max)
            {
                return input;
            }
            else
            {
                Console.WriteLine($"Пожалуйста, введите число от {min} до {max}.");
            }
        }
    }

    private string GetStringInput(string prompt)
    {
        Console.WriteLine(prompt);
        return Console.ReadLine();
    }

    private long GetLongInput(string prompt)
    {
        long value;
        while (!long.TryParse(Console.ReadLine(), out value))
        {
            Console.WriteLine("Неверный ввод. Пожалуйста, введите число.");
        }
        return value;
    }
    private decimal GetDecimalInput(string prompt)
    {
        decimal value;
        while (!decimal.TryParse(Console.ReadLine(), out value))
        {
            Console.WriteLine("Неверный ввод. Пожалуйста, введите число.");
        }
        return value;
    }

    public int GetTrackCountWithPriceGreaterThan150()
    {
        int trackCount = Database.Tracks
            .Count(t => t.Cost >= 150);

        Console.WriteLine($"Количество треков с ценой 150 и выше: {trackCount}");

        return trackCount;
    }

    public void DisplayTracksWithHighestPriceByGenre()
    {
        // Получаем все жанры
        var genres = Database.Genres.ToList();

        foreach (var genre in genres)
        {
            // Находим треки этого жанра
            var tracksInGenre = Database.Tracks
                .Where(t => t.GenreID == genre.ID)
                .ToList();

            if (!tracksInGenre.Any())
            {
                Console.WriteLine($"Жанр: {genre.Name} - Нет треков.");
                continue;
            }

            // Находим трек с максимальной ценой
            var maxPriceTrack = tracksInGenre
                .OrderByDescending(t => t.Cost)
                .First();

            // Выводим информацию о жанре, самом дорогом треке и его цене
            Console.WriteLine($"Жанр: {genre.Name, -18} | Трек: {maxPriceTrack.Name, -65} | Цена: {maxPriceTrack.Cost}");
        }
    }

    public void DisplayAlbumsAndTracksByOzzyOsbourne()
    {
        var ozzyId = Database.Artists
            .Where(a => a.Name == "Ozzy Osbourne")
            .Select(a => a.ID)
            .FirstOrDefault();

        if (ozzyId == 0)
        {
            Console.WriteLine("Исполнитель Оззи Осборн не найден.");
            return;
        }

        var albums = Database.Albums
            .Where(a => a.ArtistID == ozzyId)
            .ToList();

        if (!albums.Any())
        {
            Console.WriteLine("У Оззи Осборна нет альбомов.");
            return;
        }

        foreach (var album in albums)
        {
            Console.WriteLine($"Альбом: {album.Name}");

            // Получаем треки для каждого альбома
            var tracks = Database.Tracks
                .Where(t => t.AlbumID == album.ID)
                .Select(t => t.Name)
                .ToList();

            if (!tracks.Any())
            {
                Console.WriteLine("  Нет треков в этом альбоме.");
            }
            else
            {
                Console.WriteLine("  Треки:");
                foreach (var track in tracks)
                {
                    Console.WriteLine($"    - {track}");
                }
            }
        }
    }

    public void FindMetalArtistWithSmallestTotalTrackSize()
    {
        var result = Database.Tracks
            .Join(Database.Albums, t => t.AlbumID, a => a.ID, (t, a) => new { Track = t, Album = a })
            .Join(Database.Genres, ta => ta.Track.GenreID, g => g.ID, (ta, g) => new { Track = ta.Track, Album = ta.Album, Genre = g })
            .Where(tag => tag.Genre.Name == "Metal")
            .GroupBy(tag => tag.Album.ArtistID)
            .Select(g => new
            {
                ArtistID = g.Key,
                TotalSizeMB = g.Sum(tag => (decimal)tag.Track.Size) / (1024 * 1024)
            })
            .OrderBy(x => x.TotalSizeMB)
            .FirstOrDefault();

        if (result != null)
        {
            var artistName = Database.Artists.FirstOrDefault(a => a.ID == result.ArtistID)?.Name;
            Console.WriteLine($"Исполнитель с наименьшим суммарным размером песен в жанре Metal: {artistName} ({(int)result.TotalSizeMB} МБ)");
        }
        else
        {
            Console.WriteLine("Не найдено исполнителей в жанре Metal.");
        }
    }
}