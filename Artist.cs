public class Artist
{
    private int ID;
    private string Name;

    public int getID
    {
        get { return ID; }
        set { ID = value; }
    }

    public string getName
    {
        get { return  Name; }
        set { Name = value; }
    }
    public Artist(int id, string name)
    {
        this.getID = id;
        this.getName = name;
    }

    public override string ToString()
    {
        return $"ID: {ID, -3} | Name: {Name, -82}";
    }
}
