public class Album
{
    private int ID;
    private string Name;
    private int ArtistID;

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

    public int getArtistID
    {
        get { return ArtistID; }
        set { ArtistID = value; }
    }

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