namespace WindowsFormsApp3
{
    public class DrawingStatusModel
    {

        public int Id { get; set; }
        public string DrawingNo { get; set; }

        public string SeqNo { get; set; }
        public string RffStatus { get; set; }
        public int RffDate { get; set; }
        public string RffDescription { get; set; }
        public string Weight { get; set; }
        public double ActualWeight { get; set; }
        public bool IsImperial { get; set; }
        public string Type { get; set; }
        public string BoughtOutItem { get; set; }
        public string CatalogNumber { get; set; }
        public string BoughtOutComment { get; set; }
    }
}
