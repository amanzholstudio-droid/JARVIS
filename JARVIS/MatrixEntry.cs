namespace JARVIS
{
    // Одна строка матрицы = один вариант техники для должности.
    // У одной должности может быть несколько таких вариантов (например,
    // "системный блок ИЛИ ноутбук" — тогда это две отдельные записи).
    public class MatrixEntry
    {
        public string Location { get; set; } = "";
        public string Position { get; set; } = "";
        public string Department { get; set; } = "";
        public string Division { get; set; } = "";

        // "Действующие характеристики" — левый блок в файле
        public string CurrentType { get; set; } = "";
        public string CurrentCpu { get; set; } = "";
        public string CurrentRam { get; set; } = "";
        public string CurrentSsd { get; set; } = "";
        public string CurrentHdd { get; set; } = "";
        public string CurrentGpu { get; set; } = "";

        // "Характеристики" — правый (актуальный) блок в файле
        public string TargetType { get; set; } = "";
        public string TargetCpu { get; set; } = "";
        public string TargetRam { get; set; } = "";
        public string TargetSsd { get; set; } = "";
        public string TargetHdd { get; set; } = "";
        public string TargetGpu { get; set; } = "";

        public string Monitor { get; set; } = "";

        public bool HasMonitor => !string.IsNullOrWhiteSpace(Monitor);
        public string MonitorDisplay => HasMonitor ? "Да" : "Нет";
    }
}