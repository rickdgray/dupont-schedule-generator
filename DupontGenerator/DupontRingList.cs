namespace DupontGenerator
{
    public class DupontRingList(int firstIndex)
    {
        private readonly List<string> _dupontSchedule =
        [
            "N", "N", "N", "N", " ", " ", " ",
            "D", "D", "D", " ", "N", "N", "N",
            " ", " ", " ", "D", "D", "D", "D",
            " ", " ", " ", " ", " ", " ", " ",
            "R", "R", "R", "R", "R", " ", " "
        ];

        public string GetNext()
        {
            var value = _dupontSchedule[firstIndex];
            firstIndex = (firstIndex + 1) % _dupontSchedule.Count;
            return value;
        }
    }
}
