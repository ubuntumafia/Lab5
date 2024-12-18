using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2019.Presentation;
using System;
using System.Collections.Generic;
using Aspose.Cells;

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
        Album album = Database.Albums.Find(a => a.getID == id);
        if (album != null)
        {
            Database.Tracks.RemoveAll(t => t.getAlbumID == album.getID);
            Database.Albums.Remove(album);
            Console.WriteLine($"Альбом '{album.getName}' и все его треки удалены.");
            logger.LogAction($"Пользователь удалил альбом '{album.getName}'.");
            SaveChangesInXls();
        }
        else
        {
            Console.WriteLine($"Альбом с ID {id} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID альбома для удаления.");
        }
    }

    public void DeleteArtistById(int id)
    {
        Artist artist = Database.Artists.Find(a => a.getID == id);
        if (artist != null)
        {
            Database.Albums.RemoveAll(a => a.getArtistID == artist.getID);
            Database.Tracks.RemoveAll(t => Database.Albums.Any(a => a.getArtistID == artist.getID && a.getID == t.getAlbumID));
            Database.Artists.Remove(artist);
            Console.WriteLine($"Артист '{artist.getName}', его альбомы и треки удалены.");
            logger.LogAction($"Пользователь удалил артиста '{artist.getName}'.");
            SaveChangesInXls();
        }
        else
        {
            Console.WriteLine($"Артист с ID {id} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID артиста для удаления.");
        }
    }

    public void DeleteTrackById(int id)
    {
        Track track = Database.Tracks.Find(t => t.getID == id);
        if (track != null)
        {
            Database.Tracks.Remove(track);
            Console.WriteLine($"Трек '{track.getName}' удален.");
            logger.LogAction($"Пользователь удалил трек '{track.getName}'.");
            SaveChangesInXls();
        }
        else
        {
            Console.WriteLine($"Трек с ID {id} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID трека для удаления.");
        }
    }

    public void DeleteGenreById(int id)
    {
        Genre genre = Database.Genres.Find(g => g.getID == id);
        if (genre != null)
        {
            Database.Tracks.RemoveAll(t => t.getGenreID == genre.getID);
            Database.Genres.Remove(genre);
            Console.WriteLine($"Жанр '{genre.getName}' и все его треки удалены.");
            logger.LogAction($"Пользователь удалил жанр '{genre.getName}'.");
            SaveChangesInXls();
        }
        else
        {
            Console.WriteLine($"Жанр с ID {id} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID жанра для удаления.");
        }
    }


    public void UpdateAlbum()
    {
        int albumId = GetIntInput("Введите ID альбома для изменения:", 1, int.MaxValue);

        Album album = Database.Albums.Find(a => a.getID == albumId);
        if (album == null)
        {
            Console.WriteLine($"Альбом с ID {albumId} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID альбома для изменения.");
            return;
        }

        Console.WriteLine($"Введите новое название альбома (текущее: {album.getName}) или оставьте пустым для пропуска:");
        string newName = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(newName))
        {
            album.getName = newName;
            logger.LogAction($"Пользователь изменил название альбома {albumId} на '{album.getName}'.");
            SaveChangesInXls();
        }

        int newArtistId = GetIntInput($"Введите новый ID артиста (текущий: {album.getArtistID}) или оставьте пустым для пропуска:", 0, int.MaxValue);
        if (newArtistId != 0)
        {
            album.getArtistID = newArtistId;
            logger.LogAction($"Пользователь изменил ID артиста альбома {albumId} на '{album.getArtistID}'.");
            SaveChangesInXls();
        }

        Console.WriteLine("Альбом обновлен.");
    }

    public void UpdateArtist()
    {
        int artistId = GetIntInput("Введите ID артиста для изменения:", 1, int.MaxValue);

        Artist artist = Database.Artists.Find(a => a.getID == artistId);
        if (artist == null)
        {
            Console.WriteLine($"Артист с ID {artistId} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID артиста для изменения.");
            return;
        }

        Console.WriteLine($"Введите новое название артиста (текущее: {artist.getName}) или оставьте пустым для пропуска:");
        string newName = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(newName))
        {
            artist.getName = newName;
            logger.LogAction($"Пользователь изменил название артиста {artistId} на '{artist.getName}'.");
            SaveChangesInXls();
        }

        Console.WriteLine("Артист обновлен.");
    }

    public void UpdateTrack()
    {
        int trackId = GetIntInput("Введите ID трека для изменения:", 1, int.MaxValue);

        Track track = Database.Tracks.Find(t => t.getID == trackId);
        if (track == null)
        {
            Console.WriteLine($"Трек с ID {trackId} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID трека для изменения.");
            return;
        }

        Console.WriteLine($"Введите новое название трека (текущее: {track.getName}) или оставьте пустым для пропуска:");
        string newName = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(newName))
        {
            track.getName = newName;
            logger.LogAction($"Пользователь изменил название трека {trackId} на '{track.getName}'.");
        }

        int newAlbumId = GetIntInput($"Введите новый ID альбома (текущий: {track.getAlbumID}) или 0 для пропуска:", 0, int.MaxValue);
        if (newAlbumId != 0)
        {
            track.getAlbumID = newAlbumId;
            logger.LogAction($"Пользователь изменил ID альбома трека {trackId} на '{track.getAlbumID}'.");
        }

        int newGenreId = GetIntInput($"Введите новый ID жанра (текущий: {track.getGenreID}) или 0 для пропуска:", 0, int.MaxValue);
        if (newGenreId != 0)
        {
            track.getGenreID = newGenreId;
            logger.LogAction($"Пользователь изменил ID жанра трека {trackId} на '{track.getGenreID}'.");
        }

        Console.WriteLine($"Введите новую длительность трека (текущая: {track.getDuration}) или 0 для пропуска:");
        string input = Console.ReadLine();
        if (long.TryParse(input, out long newDuration) && newDuration != 0)
        {
            track.getDuration = newDuration;
            logger.LogAction($"Пользователь изменил продолжительность трека {trackId} на '{track.getDuration}'.");
        }
        else
        {
            Console.WriteLine("Пожалуйста, введите число!");
            logger.LogAction($"Ошибка ввода пользователем.");
        }

        int newSize = GetIntInput($"Введите новый размер трека (текущий: {track.getSize}) или 0 для пропуска:", 0, int.MaxValue);
        if (newSize != 0)
        {
            track.getSize = newSize;
            logger.LogAction($"Пользователь изменил размер трека {trackId} на '{track.getSize}'.");
        }

        Console.WriteLine($"Введите новую стоимость трека (текущая: {track.getCost}) или 0 для пропуска:");
        string input1 = Console.ReadLine();

        if (decimal.TryParse(input1, out decimal newCost) && newCost != 0)
        {
            track.getCost = newCost;
            logger.LogAction($"Пользователь изменил стоимость трека {trackId} на '{track.getCost}'.");
        }
        else
        {
            Console.WriteLine("Пожалуйста, введите число!");
            logger.LogAction($"Ошибка ввода пользователем.");
        }

        Console.WriteLine("Трек обновлен.");
        SaveChangesInXls();
    }

    public void UpdateGenre()
    {
        int genreId = GetIntInput("Введите ID жанра для изменения:", 1, int.MaxValue);

        Genre genre = Database.Genres.Find(g => g.getID == genreId);
        if (genre == null)
        {
            Console.WriteLine($"Жанр с ID {genreId} не найден.");
            logger.LogAction($"Пользователь ввел неправильный ID жанра для изменения.");
            return;
        }

        Console.WriteLine($"Введите новое название жанра (текущее: {genre.getName}) или оставьте пустым для пропуска:");
        string newName = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(newName))
        {
            genre.getName = newName;
            SaveChangesInXls();
            logger.LogAction($"Пользователь изменил название жанра {genreId} на '{genre.getName}'.");
        }

        Console.WriteLine("Жанр обновлен.");
    }


    public void AddAlbum()
    {
        string newAlbumName;
        int artistId;

        while (true)
        {
            Console.WriteLine("Введите название нового альбома:");
            newAlbumName = Console.ReadLine();

            if (string.IsNullOrEmpty(newAlbumName))
            {
                Console.WriteLine("Ошибка: название альбома не может быть пустым.");
                logger.LogAction($"Ошибка ввода названия нового альбома.");
                continue;
            }

            Console.WriteLine("Введите ID нового артиста:");
            string newArtistId = Console.ReadLine();

            if (!int.TryParse(newArtistId, out artistId))
            {
                Console.WriteLine("Ошибка: ID артиста должен быть числом.");
                logger.LogAction($"Ошибка ввода ID артиста.");
                continue;
            }

            break;
        }

        Database.Albums.Add(new Album(Database.Albums.Max(a => a.getID) + 1, newAlbumName, artistId));
        Console.WriteLine($"Альбом '{newAlbumName}' добавлен.");
        logger.LogAction($"Пользователь добавил альбом '{newAlbumName}'.");
        SaveChangesInXls();
    }

    public void AddArtist()
    {
        string newArtistName;

        while (true)
        {
            Console.WriteLine("Введите название нового артиста:");
            newArtistName = Console.ReadLine();

            if (string.IsNullOrEmpty(newArtistName))
            {
                Console.WriteLine("Ошибка: название артиста не может быть пустым.");
                logger.LogAction($"Ошибка ввода названия нового артиста.");
                continue;
            }

            break;
        }

        Database.Artists.Add(new Artist(Database.Artists.Max(a => a.getID) + 1, newArtistName));
        Console.WriteLine($"Артист '{newArtistName}' добавлен.");
        logger.LogAction($"Пользователь добавил исполнителя '{newArtistName}'.");
        SaveChangesInXls();
    }

    public void AddTrack()
    {
        string newTrackName;
        int albumId, genreId;
        long duration;
        int size;
        decimal cost;

        while (true)
        {
            Console.WriteLine("Введите название нового трека:");
            newTrackName = Console.ReadLine();

            if (string.IsNullOrEmpty(newTrackName))
            {
                Console.WriteLine("Ошибка: название трека не может быть пустым.");
                logger.LogAction($"Ошибка ввода названия нового трека.");
                continue;
            }

            Console.WriteLine("Введите ID альбома:");
            string newAlbumId = Console.ReadLine();

            if (!int.TryParse(newAlbumId, out albumId))
            {
                Console.WriteLine("Ошибка: ID альбома должен быть числом.");
                logger.LogAction($"Ошибка ввода ID альбома.");
                continue;
            }

            Console.WriteLine("Введите ID жанра:");
            string newGenreId = Console.ReadLine();

            if (!int.TryParse(newGenreId, out genreId))
            {
                Console.WriteLine("Ошибка: ID жанра должен быть числом.");
                logger.LogAction($"Ошибка ввода ID жанра.");
                continue;
            }

            Console.WriteLine("Введите длительность трека:");
            string newDuration = Console.ReadLine();

            if (!long.TryParse(newDuration, out duration))
            {
                Console.WriteLine("Ошибка: длительность трека должна быть числом.");
                logger.LogAction($"Ошибка ввода длительности трека.");
                continue;
            }

            Console.WriteLine("Введите размер трека:");
            string newSize = Console.ReadLine();

            if (!int.TryParse(newSize, out size))
            {
                Console.WriteLine("Ошибка: размер трека должен быть числом.");
                logger.LogAction($"Ошибка ввода размера трека.");
                continue;
            }

            Console.WriteLine("Введите цену трека:");
            string newCost = Console.ReadLine();

            if (!decimal.TryParse(newCost, out cost))
            {
                Console.WriteLine("Ошибка: цена трека должна быть числом.");
                logger.LogAction($"Ошибка ввода цены трека.");
                continue;
            }

            break;
        }

        Database.Tracks.Add(new Track(Database.Tracks.Max(t => t.getID) + 1, newTrackName, albumId, genreId, duration, size, cost));
        Console.WriteLine($"Трек '{newTrackName}' добавлен.");
        SaveChangesInXls();
        logger.LogAction($"Пользователь добавил трек '{newTrackName}'.");
    }


    public void AddGenre()
    {
        string newGenreName;

        while (true)
        {
            Console.WriteLine("Введите название нового жанра:");
            newGenreName = Console.ReadLine();

            if (string.IsNullOrEmpty(newGenreName))
            {
                Console.WriteLine("Ошибка: название жанра не может быть пустым.");
                logger.LogAction($"Ошибка ввода пользователем.");
                continue;
            }

            break;
        }

        Database.Genres.Add(new Genre(Database.Genres.Max(g => g.getID) + 1, newGenreName));
        Console.WriteLine($"Жанр '{newGenreName}' добавлен.");
        logger.LogAction($"Пользователь добавил жанр '{newGenreName}'.");
        SaveChangesInXls();
    }

    public void SaveChangesInXls()
    {
        try
        {
            Workbook wb = new Workbook(PathToXlsFile);
            Worksheet wsAlbums = wb.Worksheets["Альбомы"];
            Worksheet wsArtists = wb.Worksheets["Артисты"];
            Worksheet wsTracks = wb.Worksheets["Треки"];
            Worksheet wsGenres = wb.Worksheets["Жанры"];

            // Сохранение альбомов
            for (int i = 0; i < Database.Albums.Count; i++)
            {
                var album = Database.Albums[i];
                wsAlbums.Cells[$"A{i + 2}"].PutValue(album.getID);
                wsAlbums.Cells[$"B{i + 2}"].PutValue(album.getName);
                wsAlbums.Cells[$"C{i + 2}"].PutValue(album.getArtistID);
            }

            // Удаление альбомов, которые больше не существуют
            for (int i = Database.Albums.Count + 1; i <= wsAlbums.Cells.MaxDataRow + 1; i++)
            {
                wsAlbums.Cells.DeleteRow(i);
            }

            // Сохранение артистов
            for (int i = 0; i < Database.Artists.Count; i++)
            {
                var artist = Database.Artists[i];
                wsArtists.Cells[$"A{i + 2}"].PutValue(artist.getID);
                wsArtists.Cells[$"B{i + 2}"].PutValue(artist.getName);
            }

            // Удаление артистов, которые больше не существуют
            for (int i = Database.Artists.Count + 1; i <= wsArtists.Cells.MaxDataRow + 1; i++)
            {
                wsArtists.Cells.DeleteRow(i);
            }

            // Сохранение треков
            for (int i = 0; i < Database.Tracks.Count; i++)
            {
                var track = Database.Tracks[i];
                wsTracks.Cells[$"A{i + 2}"].PutValue(track.getID);
                wsTracks.Cells[$"B{i + 2}"].PutValue(track.getName);
                wsTracks.Cells[$"C{i + 2}"].PutValue(track.getAlbumID);
                wsTracks.Cells[$"D{i + 2}"].PutValue(track.getGenreID);
                wsTracks.Cells[$"E{i + 2}"].PutValue(track.getDuration);
                wsTracks.Cells[$"F{i + 2}"].PutValue(track.getSize);
                wsTracks.Cells[$"G{i + 2}"].PutValue(track.getCost);
            }

            // Удаление треков, которые больше не существуют
            for (int i = Database.Tracks.Count + 1; i <= wsTracks.Cells.MaxDataRow + 1; i++)
            {
                wsTracks.Cells.DeleteRow(i);
            }

            // Сохранение жанров
            for (int i = 0; i < Database.Genres.Count; i++)
            {
                var genre = Database.Genres[i];
                wsGenres.Cells[$"A{i + 2}"].PutValue(genre.getID);
                wsGenres.Cells[$"B{i + 2}"].PutValue(genre.getName);
            }

            // Удаление жанров, которые больше не существуют
            for (int i = Database.Genres.Count + 1; i <= wsGenres.Cells.MaxDataRow + 1; i++)
            {
                wsGenres.Cells.DeleteRow(i);
            }

            wb.Save(PathToXlsFile);
            logger.LogAction("Изменения успешно сохранены в файл.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при сохранении изменений в файл Excel: {ex.Message}");
            logger.LogAction($"Ошибка при сохранении изменений в файл Excel: {ex.Message}");
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
                logger.LogAction($"Ошибка ввода пользователем.");
            }
        }
    }

    public int GetTrackCountWithPriceGreaterThan150()
    {
        int trackCount = Database.Tracks
            .Count(t => t.getCost >= 150);

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
                .Where(t => t.getGenreID == genre.getID)
                .ToList();

            if (!tracksInGenre.Any())
            {
                Console.WriteLine($"Жанр: {genre.getName} - Нет треков.");
                continue;
            }

            // Находим трек с максимальной ценой
            var maxPriceTrack = tracksInGenre
                .OrderByDescending(t => t.getCost)
                .First();

            // Выводим информацию о жанре, самом дорогом треке и его цене
            Console.WriteLine($"Жанр: {genre.getName, -18} | Трек: {maxPriceTrack.getName, -65} | Цена: {maxPriceTrack.getCost}");
        }
    }

    public void DisplayAlbumsAndTracksByOzzyOsbourne()
    {
        var ozzyId = Database.Artists
            .Where(a => a.getName == "Ozzy Osbourne")
            .Select(a => a.getID)
            .FirstOrDefault();

        if (ozzyId == 0)
        {
            Console.WriteLine("Исполнитель Оззи Осборн не найден.");
            return;
        }

        var albums = Database.Albums
            .Where(a => a.getArtistID == ozzyId)
            .ToList();

        if (!albums.Any())
        {
            Console.WriteLine("У Оззи Осборна нет альбомов.");
            return;
        }

        foreach (var album in albums)
        {
            Console.WriteLine($"Альбом: {album.getName}");

            // Получаем треки для каждого альбома
            var tracks = Database.Tracks
                .Where(t => t.getAlbumID == album.getID)
                .Select(t => t.getName)
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
            .Join(Database.Albums, t => t.getAlbumID, a => a.getID, (t, a) => new { Track = t, Album = a })
            .Join(Database.Genres, ta => ta.Track.getGenreID, g => g.getID, (ta, g) => new { Track = ta.Track, Album = ta.Album, Genre = g })
            .Where(tag => tag.Genre.getName == "Metal")
            .GroupBy(tag => tag.Album.getArtistID)
            .Select(g => new
            {
                ArtistID = g.Key,
                TotalSizeMB = g.Sum(tag => (decimal)tag.Track.getSize) / (1024 * 1024)
            })
            .OrderBy(x => x.TotalSizeMB)
            .FirstOrDefault();

        if (result != null)
        {
            var artistName = Database.Artists.FirstOrDefault(a => a.getID == result.ArtistID)?.getName;
            Console.WriteLine($"Исполнитель с наименьшим суммарным размером песен в жанре Metal: {artistName} ({(int)result.TotalSizeMB} МБ)");
        }
        else
        {
            Console.WriteLine("Не найдено исполнителей в жанре Metal.");
        }
    }
}