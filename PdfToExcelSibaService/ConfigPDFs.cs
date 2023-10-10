
using System.Collections.Generic;

class Channel
{
    string name;
    int up_coord;
    int down_coord;
    int left_coord;
    int right_coord;

    public string Name { get => name; set => name = value; }
    public int Up_coord { get => up_coord; set => up_coord = value; }
    public int Down_coord { get => down_coord; set => down_coord = value; }
    public int Left_coord { get => left_coord; set => left_coord = value; }
    public int Right_coord { get => right_coord; set => right_coord = value; }

    public Channel(string n, int u, int d, int l, int r)
    {
        name = n;
        up_coord = u;
        down_coord = d;
        left_coord = l;
        right_coord = r;
    }
}
    class ChannelRoot
    {
        public Dictionary<string, Channel> Channels { set; get; }
    }