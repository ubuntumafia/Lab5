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
            // Удаляем все треки, связанные с этим альбомом
            Database.Tracks.RemoveAll(t => t.getAlbumID == album.getID);
            // Удаляем сам альбом
            Database.Albums.Remove(album);
            DeleteRowFromExcel("LR5-var16.xls", "Albums", id, 0);
            Console.WriteLine($"Альбом '{album.getName}' и все его треки удалены.");
            logger.LogAction($"Пользователь удалил альбом '{album.getName}'.");
        }
        else
        {
            Console.WriteLine($"Альбом с ID {id} не найден.");
            logger.LogAction("Пользователь пытался удалить альбом.");
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
            DeleteRowFromExcel("LR5-var16.xls", "Artists", id, 1);
            Console.WriteLine($"Артист '{artist.getName}', его альбомы и треки удалены.");
            logger.LogAction($"Пользователь удалил артиста '{artist.getName}'.");
        }
        else
        {
            Console.WriteLine($"Артист с ID {id} не найден.");
            logger.LogAction($"Пользователь пытался удалить артиста.");
        }
    }

    public void DeleteTrackById(int id)
    {
        Track track = Database.Tracks.Find(t => t.getID == id);
        if (track != null)
        {
            Database.Tracks.Remove(track);
            DeleteRowFromExcel("LR5-var16.xls", "Tracks", id, 2);
            Console.WriteLine($"Трек '{track.getName}' удален.");
            logger.LogAction($"Пользователь удалил трек '{track.getName}'.");
        }
        else
        {
            Console.WriteLine($"Трек с ID {id} не найден.");
            logger.LogAction($"Пользователь пытался удалить артиста.");
        }
    }

    public void DeleteGenreById(int id)
    {
        Genre genre = Database.Genres.Find(g => g.getID == id);
        if (genre != null)
        {
            Database.Tracks.RemoveAll(t => t.getGenreID == genre.getID);
            Database.Genres.Remove(genre);
            DeleteRowFromExcel("LR5-var16.xls", "Genres", id, 3);
            Console.WriteLine($"Жанр '{genre.getName}' и все его треки удалены.");
            logger.LogAction($"Пользователь удалил жанр '{genre.getName}'.");
        }
        else
        {
            Console.WriteLine($"Жанр с ID {id} не найден.");
            logger.LogAction($"Пользователь пытался удалить жанр.");
        }
    }

    public static void DeleteRowFromExcel(string filePath, string sheetName, int recordId, int idColumnIndex)
    {
        var workbook = new Workbook(filePath);

        var worksheet = workbook.Worksheets[sheetName];
        if (worksheet == null) 
        {
            Console.WriteLine($"Лист с именем '{sheetName}' не найден.");
            return;
        }

        // Получаем ячейки на листе
        var cells = worksheet.Cells;
        int rowToDelete = -1;

        // Перебираем строки на листе
        for (int i = 0; i <= cells.MaxDataRow; i++)
        {
            var cellValue = cells[i, idColumnIndex].StringValue;
            if (int.TryParse(cellValue, out int currentId) && currentId == recordId) // Если ID совпадает
            {
                rowToDelete = i; // Запоминаем индекс строки
                break;
            }
        }

        if (rowToDelete >= 0)
        {
            cells.DeleteRow(rowToDelete);
            workbook.Save(filePath);
            Console.WriteLine($"Запись с ID {recordId} удалена из листа '{sheetName}'.");
        }
        else
        {
            Console.WriteLine($"Запись с ID {recordId} не найдена на листе '{sheetName}'.");
        }
    }


    public void UpdateAlbum()
    {
        int albumId = GetIntInput("Введите ID альбома для изменения:", 1, int.MaxValue);

        Album album = Database.Albums.Find(a => a.getID == albumId);
        if (album == null)
        {
            Console.WriteLine($"Альбом с ID {albumId} не найден.");
            logger.LogAction($"Пользователь пытался изменить альбом.");
            return;
        }

        string newName = GetStringInput($"Введите новое название альбома (текущее: {album.getName}) или 0 для пропуска:");
        if (newName != "0")
        {
            album.getName = newName;
            logger.LogAction($"Пользователь изменил название альбома {albumId} на '{album.getName}'.");
        }

        int newArtistId = GetIntInput($"Введите новый ID артиста (текущий: {album.getArtistID}) или 0 для пропуска:", 0, int.MaxValue);
        if (newArtistId != 0)
        {
            album.getArtistID = newArtistId;
            logger.LogAction($"Пользователь изменил ID артиста альбома {albumId} на '{album.getArtistID}'.");
        }

        Console.WriteLine("Альбом  обновлен.");
    }

    public void UpdateArtist()
    {
        int artistId = GetIntInput("Введите ID артиста для изменения:", 1, int.MaxValue);

        Artist artist = Database.Artists.Find(a => a.getID == artistId);
        if (artist == null)
        {
            Console.WriteLine($"Артист с ID {artistId} не найден.");
            logger.LogAction($"Пользователь пытался изменить артиста.");
            return;
        }

        string newName = GetStringInput($"Введите новое название артиста (текущее: {artist.getName}) или 0 для пропуска:");
        if (newName != "0")
        {
            artist.getName = newName;
            logger.LogAction($"Пользователь изменил название артиста {artistId} на '{artist.getName}'.");
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
            logger.LogAction($"Пользователь пытался изменить трек.");
            return;
        }

        string newName = GetStringInput($"Введите новое название трека (текущее: {track.getName}) или 0 для пропуска:");
        if (newName != "0")
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

        long newDuration = GetLongInput($"Введите новую длительность трека (текущая: {track.getDuration}) или 0 для пропуска:");
        if (newDuration != 0)
        {
            track.getDuration = newDuration;
            logger.LogAction($"Пользователь изменил продолжительность трека {trackId} на '{track.getDuration}'.");
        }

        int newSize = GetIntInput($"Введите новый размер трека (текущий: {track.getSize}) или 0 для пропуска:", 0, int.MaxValue);
        if (newSize != 0)
        {
            track.getSize = newSize;
            logger.LogAction($"Пользователь изменил размер трека {trackId} на '{track.getSize}'.");
        }

        decimal newCost = GetDecimalInput($"Введите новую стоимость трека (текущая: {track.getCost}) или 0 для пропуска:");
        if (newCost != 0)
        {
            track.getCost = newCost;
            logger.LogAction($"Пользователь изменил стоимость трека {trackId} на '{track.getCost}'.");
        }

        Console.WriteLine("Трек обновлен.");
    }

    public void UpdateGenre()
    {
        int genreId = GetIntInput("Введите ID жанра для изменения:", 1, int.MaxValue);

        Genre genre = Database.Genres.Find(g => g.getID == genreId);
        if (genre == null)
        {
            Console.WriteLine($"Жанр с ID {genreId} не найден.");
            logger.LogAction($"Пользователь пытался изменить жанр.");
            return;
        }

        string newName = GetStringInput($"Введите новое название жанра (текущее: {genre.getName}) или 0 для пропуска:");
        if (newName != "0")
        {
            genre.getName = newName;
            logger.LogAction($"Пользователь изменил название жанра {genreId} на '{genre.getName}'.");
        }

        Console.WriteLine($"Жанр обновлен.");
    }

    public static void UpdateRowInExcel(string filePath, string sheetName, int recordId, params object[] newValues)
    {
        var workbook = new Workbook(filePath);

        var worksheet = workbook.Worksheets[sheetName];
        if (worksheet == null)
        {
            Console.WriteLine($"Лист с именем '{sheetName}' не найден.");
            return;
        }

        var cells = worksheet.Cells;
        int rowToUpdate = -1;

        // Перебираем все строки в листе
        for (int i = 0; i <= cells.MaxDataRow; i++)
        {
            // Считываем значение из первой колонки строки (предполагается, что ID находится в колонке 0)
            var cellValue = cells[i, 0].StringValue;
            if (int.TryParse(cellValue, out int currentId) && currentId == recordId)
            {
                rowToUpdate = i;
                break;
            }
        }

        if (rowToUpdate >= 0)
        {
            if (newValues.Length > 0)
            {
                cells[rowToUpdate, 1].PutValue(newValues[0]);
            }

            if (newValues.Length > 1)
            {
                for (int col = 2; col < newValues.Length + 1; col++)
                {
                    cells[rowToUpdate, col].PutValue(newValues[col - 1]);
                }
            }

            workbook.Save(filePath);
            Console.WriteLine($"Запись с ID {recordId} обновлена на листе '{sheetName}'.");
        }
        else
        {
            // Если строка с указанным ID не найдена, выводим сообщение
            Console.WriteLine($"Запись с ID {recordId} не найдена на листе '{sheetName}'.");
        }
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
            Database.Albums.Add(new Album(Database.Albums.Max(a => a.getID) + 1, newAlbumName, artistId));
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
            Database.Artists.Add(new Artist(Database.Artists.Max(a => a.getID) + 1, newArtistName));
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

            Database.Tracks.Add(new Track(Database.Tracks.Max(t => t.getID) + 1, newTrackName, albumId, genreId, duration, size, cost));
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
            Database.Genres.Add(new Genre(Database.Genres.Max(g => g.getID) + 1, newGenreName));
            Console.WriteLine($"Жанр '{newGenreName}' добавлен.");
            logger.LogAction($"Пользователь добавил жанр '{newGenreName}'.");
        }
    }

    public static void AddRowToExcel(string filePath, string sheetName, int idColumnIndex, object idValue, params object[] values)
    {
        var workbook = new Workbook(filePath);

        // Получаем нужный лист по имени
        var worksheet = workbook.Worksheets[sheetName];
        if (worksheet == null)
        {
            Console.WriteLine($"Лист с именем '{sheetName}' не найден.");
            return;
        }

        bool idExists = false;
        for (int row = 0; row <= worksheet.Cells.MaxDataRow; row++)
        {
            if (worksheet.Cells[row, idColumnIndex].Value != null && worksheet.Cells[row, idColumnIndex].Value.ToString() == idValue.ToString())
            {
                idExists = true;
                break;
            }

        }

        if (idExists)
        {
            Console.WriteLine($"Запись с ID '{idValue}' уже существует в листе '{sheetName}'.");
            return;
        }

        int newRow = worksheet.Cells.MaxDataRow + 1;

        for (int col = 0; col < values.Length; col++)
        {
            worksheet.Cells[newRow, col].PutValue(values[col]);
        }

        workbook.Save(filePath);

        Console.WriteLine($"Добавлена запись в лист '{sheetName}'.");
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