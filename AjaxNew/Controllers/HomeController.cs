using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AjaxNew.Controllers
{
    public class HomeController : Controller
    {
        public class SearchResultLine
        {
            public string Fio { get; set; }
            public string PhoneNumber { get; set; }
            public int? Room { get; set; }
            public string RoomClass { get; set; }
            public int? ID { get; set; }
        }
        public ActionResult Index(string pattern)
        {
            List<SearchResultLine> result;
            using (DBMod_Hotel db = new DBMod_Hotel())
            {
                result = db.Clients
                    .Join
                    (
                       db.RoomClient,
                       s => s.ID,
                       sc => sc.ID_CLIENT,
                       (s, sc) => new
                       {
                           Name = s.FIO,
                           Number = s.PHONE_NUMBER,
                           id_s_c = sc.ID,
                           id_c = sc.ID_ROOM
                       }
                    ).Join
                    (
                    db.Rooms,
                    sc => sc.id_c,
                    c => c.ID,
                    (sc, c) => new SearchResultLine()
                    {
                        Fio = sc.Name,
                        PhoneNumber = sc.Number,
                        Room = sc.id_c,
                        RoomClass = c.CLASS,
                        ID = sc.id_s_c
                    }
                    ).ToList();
            }
            if(pattern == null)
            {
                ViewBag.SearchData = result;
                return View();
            }
            else
            {
                result = result.Where((p) => p.Fio.Contains(pattern)).ToList();
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }
        public MemoryStream GenerateWord(string[,] data)
        {
            MemoryStream mStream = new MemoryStream();
            WordprocessingDocument document = WordprocessingDocument.Create(mStream, WordprocessingDocumentType.Document, true);

            MainDocumentPart mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            Table table = new Table();
            body.AppendChild(table);

            TableProperties props = new TableProperties
            (
                new TableBorders
                (
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 }
                )
            );

            table.AppendChild<TableProperties>(props);

            for (var i = 0; i < data.GetUpperBound(0); i++)
            {
                var tr = new TableRow();
                for (var j = 0; j <= data.GetUpperBound(1); j++)
                {
                    var tc = new TableCell();
                    tr.Append(new Paragraph(new Run(new Text(data[i, j]))));

                    tr.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                    tr.Append(tc);
                }
                table.Append(tr);
            }
            mainPart.Document.Save();
            document.Clone();
            mStream.Position = 0;
            return mStream;
        }
        public FileStreamResult GetWord()
        {
            List<SearchResultLine> result;
            using (DBMod_Hotel db = new DBMod_Hotel())
            {
                result = db.Clients
                    .Join
                    (
                       db.RoomClient,
                       s => s.ID,
                       sc => sc.ID_CLIENT,
                       (s, sc) => new
                       {
                           Name = s.FIO,
                           Number = s.PHONE_NUMBER,
                           id_s_c = sc.ID,
                           id_c = sc.ID_ROOM
                       }
                    ).Join
                    (
                    db.Rooms,
                    sc => sc.id_c,
                    c => c.ID,
                    (sc, c) => new SearchResultLine()
                    {
                        Fio = sc.Name,
                        PhoneNumber = sc.Number,
                        Room = sc.id_c,
                        RoomClass = c.CLASS,
                        ID = sc.id_s_c
                    }
                    ).ToList();
            }
            string[,] data = new string[result.Count + 1, 5];
            for (int i = 0; i < result.Count; i++)
            { 
                data[i, 0] = result[i].Fio;
                data[i, 1] = result[i].PhoneNumber;
                data[i, 2] = result[i].Room.ToString();
                data[i, 3] = result[i].RoomClass;
                data[i, 4] = result[i].ID.ToString();
            }
            MemoryStream memoryStream = GenerateWord(data);
            return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = "demo.docx"
            };
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public void DeleteRow(int delId)
        {
            DBMod_Hotel db = new DBMod_Hotel();
            var dat = db.RoomClient.Where(x => x.ID == delId).FirstOrDefault();
            db.RoomClient.Remove(dat);
            db.SaveChanges();
        }
        public void AddRow(int roomId, int clientId)
        {
            DBMod_Hotel db = new DBMod_Hotel();
            db.RoomClient.Add(new RoomClient() { ID_CLIENT = clientId, ID_ROOM = roomId });
            db.SaveChanges();
        }
    }
}