using System;
using System.Collections.Generic;
using Aspose.Cells;

public class Database
{
    public List<Album> Albums;
    public List<Artist> Artists;
    public List<Track> Tracks;
    public List<Genre> Genres;

    public Database(string pathToXlsFile)
    {
        Albums = new List<Album>();
        Artists = new List<Artist>();
        Tracks = new List<Track>();
        Genres = new List<Genre>();

        LoadDataFromExcel(pathToXlsFile);
    }

    public void LoadDataFromExcel(string pathToXlsFile)
    {
        Workbook workbook = new Workbook(pathToXlsFile);
        WorksheetCollection worksheets = workbook.Worksheets;

        LoadAlbumsFromWorksheet(worksheets[0]);
        LoadArtistsFromWorksheet(worksheets[1]);
        LoadTracksFromWorksheet(worksheets[2]);
        LoadGenresFromWorksheet(worksheets[3]);
    }

    private void LoadAlbumsFromWorksheet(Worksheet worksheet)
    {
        for (int i = 1; i <= worksheet.Cells.MaxRow; i++)
        {
            var cellValue = worksheet.Cells[i, 0].Value;
            if (cellValue != null && int.TryParse(cellValue.ToString(), out int id))
            {
                string name = worksheet.Cells[i, 1].Value?.ToString() ?? "";
                int artistId = int.Parse(worksheet.Cells[i, 2].Value?.ToString() ?? "0");
                Albums.Add(new Album(id, name, artistId));
            }
        }
    }

    private void LoadArtistsFromWorksheet(Worksheet worksheet)
    {
        for (int i = 1; i <= worksheet.Cells.MaxRow; i++)
        {
            var cellValue = worksheet.Cells[i, 0].Value;
            if (cellValue != null && int.TryParse(cellValue.ToString(), out int id))
            {
                string name = worksheet.Cells[i, 1].Value?.ToString() ?? "";
                Artists.Add(new Artist(id, name));
            }
        }
    }

    private void LoadTracksFromWorksheet(Worksheet worksheet)
    {
        for (int i = 1; i <= worksheet.Cells.MaxRow; i++)
        {
            var cellValue = worksheet.Cells[i, 0].Value;
            if (cellValue != null && int.TryParse(cellValue.ToString(), out int id))
            {
                string name = worksheet.Cells[i, 1].Value?.ToString() ?? "";
                int albumId = int.Parse(worksheet.Cells[i, 2].Value?.ToString() ?? "0");
                int genreId = int.Parse(worksheet.Cells[i, 3].Value?.ToString() ?? "0");
                int duration = int.Parse(worksheet.Cells[i, 4].Value?.ToString() ?? "0");
                int size = int.Parse(worksheet.Cells[i, 5].Value?.ToString() ?? "0");
                decimal cost = decimal.Parse(worksheet.Cells[i, 6].Value?.ToString() ?? "0");
                Tracks.Add(new Track (id, name, albumId, genreId, duration, size, cost));
            }
        }
    }

    private void LoadGenresFromWorksheet(Worksheet worksheet)
    {
        for (int i = 1; i <= worksheet.Cells.MaxRow; i++)
        {
            var cellValue = worksheet.Cells[i, 0].Value;
            if (cellValue != null && int.TryParse(cellValue.ToString(), out int id))
            {
                string name = worksheet.Cells[i, 1].Value?.ToString() ?? "";
                Genres.Add(new Genre(id,name));
            }
        }
    }
}