public class Album
{
    public int ID { get; set; }
    public string Name { get; set; }
    public int ArtistID { get; set; }

    public Album(int id, string name, int artistID)
    {
        ID = id;
        Name = name;
        ArtistID = artistID;
    }

    public override string ToString()
    {
        return $"ID: {ID, -3} | Name: {Name, -95} | ArtistID: {ArtistID, -3}";
    }
}