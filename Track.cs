public class Track
{
    public int ID { get; set; }
    public string Name { get; set; }
    public int AlbumID { get; set; }
    public int GenreID { get; set; }
    public long Duration { get; set; }
    public long Size { get; set; }
    public decimal Cost { get; set; }

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