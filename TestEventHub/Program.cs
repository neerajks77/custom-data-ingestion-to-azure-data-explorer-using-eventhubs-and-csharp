using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.ServiceBus.Messaging;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace TestEventHub
{
    class Program
    {
        private static readonly string filePath = Environment.CurrentDirectory + @"\Sampledata\StormEvents.csv";
        private static _Application excel = new _Excel.Application();
        private static Workbook wb = excel.Workbooks.Open(filePath);
        private static Worksheet ws = wb.ActiveSheet;
        private static _Excel.Range range = ws.UsedRange;
        private static readonly int rangeRows = range.Rows.Count;
        private static readonly int rangeColumns = range.Columns.Count;

        private static EventHubClient eventHubClient;
        private const string EventHubConnectionString = "Endpoint=sb://aqtcsleh.servicebus.windows.net/;SharedAccessKeyName=RootManageSharedAccessKey;SharedAccessKey=dEIL5mTgv7YqmaWkxFoEzX2pc1avU86Mg8R94hhn/io=";
        private const string EventHubName = "aqtcsleh";

        static void Main(string[] args)
        {
            Console.WriteLine("Press Ctrl-C to stop the sender process");
            Console.WriteLine("Press Enter to start now");
            Console.ReadLine();
            SendingRandonMessages();
        }

        private static void SendingRandonMessages()
        {
            eventHubClient = EventHubClient.CreateFromConnectionString(EventHubConnectionString, EventHubName);
            DateTime StartTime, EndTime;
            int EpisodeId, EventId, InjuriesDirect, InjuriesIndirect, DeathsDirect, DeathsIndirect, DamageProperty, DamageCrops;
            string State, EventType, Source, BeginLocation, EndLocation, EventNarrative, EpisodeNarrative;
            decimal BeginLat, BeginLon, EndLat, EndLon;
            dynamic StormSummary;

            for (int row = 2; row < rangeRows; row++)
            {
                int col = 1;
                StartTime = DateTime.FromOADate((ws.Cells[row, col] as _Excel.Range).Value);
                EndTime = DateTime.FromOADate((ws.Cells[row, col + 1] as _Excel.Range).Value);
                EpisodeId = Convert.ToInt32((ws.Cells[row, col + 2] as _Excel.Range).Value);
                EventId = Convert.ToInt32((ws.Cells[row, col + 3] as _Excel.Range).Value);
                State = Convert.ToString((ws.Cells[row, col + 4] as _Excel.Range).Value);
                EventType = Convert.ToString((ws.Cells[row, col + 5] as _Excel.Range).Value);
                InjuriesDirect = Convert.ToInt32((ws.Cells[row, col + 6] as _Excel.Range).Value);
                InjuriesIndirect = Convert.ToInt32((ws.Cells[row, col + 7] as _Excel.Range).Value);
                DeathsDirect = Convert.ToInt32((ws.Cells[row, col + 8] as _Excel.Range).Value);
                DeathsIndirect = Convert.ToInt32((ws.Cells[row, col + 9] as _Excel.Range).Value);
                DamageProperty = Convert.ToInt32((ws.Cells[row, col + 10] as _Excel.Range).Value);
                DamageCrops = Convert.ToInt32((ws.Cells[row, col + 11] as _Excel.Range).Value);
                Source = Convert.ToString((ws.Cells[row, col + 12] as _Excel.Range).Value);
                BeginLocation = Convert.ToString((ws.Cells[row, col + 13] as _Excel.Range).Value);
                EndLocation = Convert.ToString((ws.Cells[row, col + 14] as _Excel.Range).Value);
                BeginLat = Convert.ToDecimal((ws.Cells[row, col + 15] as _Excel.Range).Value);
                BeginLon = Convert.ToDecimal((ws.Cells[row, col + 16] as _Excel.Range).Value);
                EndLat = Convert.ToDecimal((ws.Cells[row, col + 17] as _Excel.Range).Value);
                EndLon = Convert.ToDecimal((ws.Cells[row, col + 18] as _Excel.Range).Value);
                EpisodeNarrative = Convert.ToString((ws.Cells[row, col + 19] as _Excel.Range).Value);
                EventNarrative = Convert.ToString((ws.Cells[row, col + 20] as _Excel.Range).Value);
                object summary = (ws.Cells[row, col + 21] as _Excel.Range).Value;
                StormSummary = Convert.ChangeType((ws.Cells[row, col + 21] as _Excel.Range).Value, summary.GetType());
                try
                {

                    List<string> list = Enumerable
                        .Range(0, 1)
                        .Select(recordNumber => $"{{\"StartTime\": \"{StartTime}\", \"EndTime\": \"{EndTime}\", \"EpisodeId\": {EpisodeId}, \"EventId\": {EventId}, \"State\": \"{State}\", \"EventType\": \"{EventType}\", \"InjuriesDirect\": {InjuriesDirect}, \"InjuriesIndirect\": {InjuriesIndirect}, \"DeathsDirect\": {DeathsDirect}, \"DeathsIndirect\": {DeathsIndirect}, \"DamageProperty\": {DamageProperty}, \"DamageCrops\": {DamageCrops}, \"Source\": \"{Source}\", \"BeginLocation\": \"{BeginLocation}\", \"EndLocation\": \"{EndLocation}\", \"BeginLat\": {BeginLat}, \"BeginLon\": {BeginLon}, \"EndLat\": {EndLat}, \"EndLon\": {EndLon}, \"EpisodeNarrative\": \"{EpisodeNarrative}\", \"EventNarrative\": \"{EventNarrative}\", \"StormSummary\": {StormSummary}}}")
                        .ToList();

                    string recordMessage = string.Join(Environment.NewLine, list);
                    EventData eventData = new EventData(Encoding.UTF8.GetBytes(recordMessage));
                    eventHubClient.SendAsync(eventData);
                    Console.WriteLine("Row - {0} sent to EventHub at {1}", recordMessage, DateTime.Now);
                }
                catch (Exception Ex)
                {
                    Console.WriteLine("Row - {0} could not be sent to EventHub because of the exception {1}", row, Ex.Message.ToString());
                }
                Thread.Sleep(1000);
            }
        }
    }
}
