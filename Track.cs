public class Track
{
    private int ID;
    private string Name;
    private int AlbumID;
    private int GenreID;
    private long Duration;
    private long Size;
    private decimal Cost;

    public int getID
    {
        get { return ID; }
        set { ID = value; }
    }

    public string getName
    {
        get { return Name; }
        set { Name = value; }
    }

    public int getAlbumID
    {
        get { return AlbumID; }
        set { AlbumID = value; }
    }

    public int getGenreID
    {
        get { return GenreID; }
        set { GenreID = value; }
    }

    public long getDuration
    {
        get { return Duration; }
        set { Duration = value; }
    }

    public long getSize
    {
        get { return Size; }
        set { Size = value; }
    }

    public decimal getCost
    {
        get { return Cost; }
        set { Cost = value; }
    }

    public Track(int id, string name, int albumID, int genreID, long duration, long size, decimal cost)
    {
        ID = id;
        Name = name;
        AlbumID = albumID;
        GenreID = genreID;
        Duration = duration;
        Size = size;
        Cost = cost;
    }

    public override string ToString()
    {
        return $"ID: {ID, -4} | Name: {Name, -109} | AlbumID: {AlbumID, -3} | GenreID: {GenreID, -2} | Duration: {Duration, -7}ms | Size: {Size, -10}B | Cost: {Cost, -3}P";
    }
}