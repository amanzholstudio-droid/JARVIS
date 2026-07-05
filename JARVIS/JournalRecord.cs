namespace JARVIS
{
    // Одна строка из журнала (journal.xlsx) — используется только для поиска/отображения,
    // не для записи (запись остаётся в SaveToJournal через ClosedXML напрямую).
    public class JournalRecord
    {
        public string DateTime { get; set; } = "";
        public string OperationType { get; set; } = "";
        public string Basis { get; set; } = "";
        public string IncidentNo { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string EmployeePosition { get; set; } = "";
        public string EmployeeIIN { get; set; } = "";
        public string ToName { get; set; } = "";
        public string ToPosition { get; set; } = "";
        public string ToIIN { get; set; } = "";
        public string DeviceType { get; set; } = "";
        public string Inv { get; set; } = "";
        public string SN { get; set; } = "";
    }
}