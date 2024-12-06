public class Genre
{
    public int ID { get; set; }
    public string Name { get; set; }

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