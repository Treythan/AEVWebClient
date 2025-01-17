using System.Text.Json;

namespace AEVWebClient.Models
{
    public class ScheduledUnit
    {
        public DateTime? StartDate { get; set; }
        public DateTime? ProjectedDeliveryDate { get; set; }
        public string StartPoint { get; set; }
        public string ValueStream { get; set; }
        public string WorkOrder { get; set; }
        public string JobNumber { get; set; }
        public string Customer { get; set; }
        public string Box { get; set; }
        public string Chassis { get; set; }
        public string Indicator { get; set; }
        public bool? Complete { get; set; }
        public DateTime? FirstDayOfProdWeek { get; set; }
        public string DayAndNumber { get; set; }
        public string LineOrder { get; set; }
        public string BuildNumber { get; set; }

        public override string ToString()
        {
            return JsonSerializer.Serialize(this);
        }
    }
}
