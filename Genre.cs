public class Genre
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
        get { return Name; }
        set { Name = value; }
    }

    public Genre(int id, string name)
    {
        ID = id;
        Name = name;
    }

    public override string ToString()
    {
        return $"ID: {ID, -2} | Name: {Name, -18}";
    }
}