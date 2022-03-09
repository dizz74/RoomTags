using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomTags
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        private Document doc;
         
        private RoomTagType roomTagType;
        private List<Level> allLevelsOrdered;
        private Level baseLevel1;
        private int offset;
        private bool placeAllProject;


        /** 
         * Расстановка марок помещений вида "1_23" во всех помещениях.
         */
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDocument = uiApp.ActiveUIDocument;
            doc = uiDocument.Document;

            LevelsPrepare(); //подготовка уровней, узнаем базовый уровень 0.00мм и offset - для расчета этажности



            TaskDialog td = new TaskDialog("Выберите");
            td.MainContent = "Где выполнить действие?";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                               "На текущем виде",
                               "На текущем виде(вид должен быть план)");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                                "На всех планах",
                                "На всех планах проекта");
            
            switch (td.Show())
            {
                case TaskDialogResult.CommandLink1:
                    placeAllProject = false;
                    break;

                case TaskDialogResult.CommandLink2:
                    placeAllProject = true;
                    break;
            }

            if (!placeAllProject) {
                //если обрабатываем только текущий вид, то проверяем что активный вид - ViewPlan?
                View view = doc.ActiveView;
                if (view == null || !(view is ViewPlan))
                {
                    message = "Текущий вид - не план.\nВыберите план!";
                    return Result.Failed;
                }
            
            } 
            

            roomTagType = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_RoomTags).
               OfType<RoomTagType>()
               .Where(rt => rt.Name.Equals("Марка Этаж_номер"))
               .FirstOrDefault();

            if (roomTagType == null)
            {
                message = "Марка \"Марка Этаж_номер\" не загружена!";
                return Result.Failed;
            }

            AskPlaceRooms();

            PlaceTags();



            return Result.Succeeded;
        }

        private void PlaceTags()
        {
            int total = 0;

            if (placeAllProject)
            {
                //расставляем во всем проекте на всех планах
                new FilteredElementCollector(doc).OfClass(typeof(ViewPlan))
                     .OfType<ViewPlan>()
                     .Where(x => !x.IsTemplate)
                     .ToList()
                     .ForEach(vp => total += PlaceTagsPerViewPlan(vp));

            }
            else
            {
                //расставляем в текущем виде 
                total = PlaceTagsPerViewPlan(doc.ActiveView as ViewPlan);
            }


            TaskDialog.Show("Установка меток завершена", $"Установлено  меток  = {total}");
        }

        private void AskPlaceRooms()
        {
            int total = 0; 
            TaskDialog td = new TaskDialog("Выберите");
            td.MainContent = "Расставить помещения автоматически?";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                               "Нет");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                                "Да");
            switch (td.Show())
            {
              
                case TaskDialogResult.CommandLink1:
                    return;
                case TaskDialogResult.CommandLink2:
                    //расставляем помещения
                    if (placeAllProject)
                    {
                        foreach (Level level in allLevelsOrdered)
                        {
                            total += PlaceRooms(level);
                        }
                    }
                    else
                    {

                        ViewPlan viewPlan = doc.ActiveView as ViewPlan;
                        Level currentLvl = GetLevelByName(viewPlan.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsString());
                        total += PlaceRooms(currentLvl);
                    }
                    TaskDialog.Show("Установка помещений завершена", $"Установлено  помещений  = {total}");
                    break;
                
            }
             

            
        }

        /**
         * Расстановка помещений на уровне
         * */
        private int PlaceRooms(Level level)
        {
            int x=0;
            using (var ts = new Transaction(doc, $"установка помещений на этаже ({level.Name})"))
            {
                ts.Start();
                PhaseArray phases = doc.Phases;
                Phase createRoomsInPhase = phases.get_Item(phases.Size - 1);
                PlanTopology topology = doc.get_PlanTopology(level, createRoomsInPhase);
                PlanCircuitSet circuitSet = topology.Circuits;
                 
                foreach (PlanCircuit circuit in circuitSet)
                {
                    if (circuit == null) continue;
                    if (circuit.Area < 10) continue;//игнорируем вент 
                    if (!circuit.IsRoomLocated)
                    {
                        try
                        {
                            Room room = doc.Create.NewRoom(null, circuit);
                            room.Name = x.ToString();
                            x++;
                        }
                        catch (Exception e) { 
                        }
                    }
                }
                ts.Commit();
            }
            return x;
        }

        /*
         * заполняем все уровни, базовый уровень, offset для отсчета
        */
        private void LevelsPrepare()
        {
            allLevelsOrdered = new FilteredElementCollector(doc)
                                              .OfClass(typeof(Level))
                                              .OfType<Level>()
                                              .OrderBy(x => x.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())
                                              .ToList();
            baseLevel1 = allLevelsOrdered.Where(x => x.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble() == 0).FirstOrDefault();//1 этаж, уровень 0.00мм

            if (baseLevel1 == null)
            {
                //уровень с base=0.00 не найден, используем этаж "Этаж 1"/"Level 1"
                baseLevel1 = GetLevelByName(" 1", true);
            }
            offset = 0;
            for (int i = 0; i < allLevelsOrdered.Count; i++)
            {
                if (allLevelsOrdered[i].Id.Equals(baseLevel1.Id))
                {
                    offset = i;
                    TaskDialog.Show("Базовый уровень", "За базовый 1 этаж принимаем " + allLevelsOrdered[i].Name);
                    break;
                }
            }
        }

        /*
         Размещение меток на ViewPlan
         */
        private int PlaceTagsPerViewPlan(ViewPlan viewPlan)
        {

            Level currentLvl = GetLevelByName(viewPlan.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsString());


            List<Room> rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Room>()
                .Where(r => r.LevelId.Equals(currentLvl.Id))
                .ToList();

            if (rooms.Count == 0)
            {                
                //TaskDialog.Show("Предупреждение", $"На плане {viewPlan.Name} - 0 помещений :("); 
                return 0;
            }

            string nLevel = getFloorLevel(currentLvl);//"1"
            int tagsPlaced = 0;
            using (var ts = new Transaction(doc, $"установка меток на плане {viewPlan.Name}"))
            {
                ts.Start();
                foreach (Room room in rooms)
                {
                    if (PlaceMyTag(viewPlan,room, nLevel)) tagsPlaced++;
                }
                ts.Commit();
            }
            return tagsPlaced;
        }

        /*
        получаем этажность уровня. Этаж 1 => 1; подвал => 0.
         */
        private string getFloorLevel(Level level)
        {
           
             
            int totalLevels = allLevelsOrdered.Count();
          

            for (int i = 0; i < totalLevels; i++)
            {
                if (allLevelsOrdered[i].Id.Equals(level.Id)) return (i - offset + 1).ToString();
            }

            //что-то не так, возвращаем имя
            return level.Name;

        }
                                      

        private bool PlaceMyTag(ViewPlan viewPlan,Room room, string floor_level)
        {

            Parameter nLevel = room.LookupParameter("Номер_этажа");
            nLevel?.Set(floor_level);//присваиваем помещению номер этажа
            if (room.Location == null)
            {
                //автоматически созданные помещения некоторые имеют Location null,видимо ошибочно созданные, пропускаем
                return false;
            }
            XYZ point = (room.Location as LocationPoint).Point;
            UV centre = new UV(point.X, point.Y);
            RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), centre, viewPlan.Id);
            roomTag.RoomTagType = roomTagType;
            return true;
        }


        private Level GetLevelByName(string levelName, bool endName = false)
        {

            return new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Level))
                                                 .OfType<Level>()
                                                 .Where(l =>
                                                 endName ? l.Name.EndsWith(levelName) : l.Name.Equals(levelName)).FirstOrDefault();
        }

    }

}
