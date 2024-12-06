public class Artist
{
    public int ID { get; set; }
    public string Name { get; set; }

    public Artist(int id, string name)
    {
        ID = id;
        Name = name;
    }

    public override string ToString()
    {
        return $"ID: {ID, -3} | Name: {Name, -82}";
    }
}
