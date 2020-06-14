using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace NIPIGAS_BomExtractorLogic
{
    public class BOM_Extractor 
    {
        public string FilePath { get; set; } //Параметр, который определяет путь к целевому PDF

        public int PageID { get; set; } //Параметр, определяющий номер страницы в целевом PDF. Так же подаётся извне

        //По умолчанию эти 2 списка имеют значение new List<string>() - то есть 2 инициализированных, пустых списка. Будем добавлять в них строки по ходу действия
        private List<string> Components { get; set; } = new List<string>(); //Список строк, извлечённых из области компонентов
        private List<string> Parameters { get; set; } = new List<string>(); //Список строк, извлечённых из области параметров

        //Таблица, в которую мы будет записывать результат. По умолчанию пустая
        public DataTable BomDataTable { get; set; } = new DataTable();

        //Словарь значений извлеченных областей текста: Чертёж / Номер ревизии и т.п.
        public Dictionary<string, string> DrawingElements { get; set; } = new Dictionary<string, string>();

        //Словарь строк и массивов флоатов. Нужен для определения координат таких областей как "Чертёж", "Номер ревизии", и т.п. по умолчанию пустой
        private Dictionary<string, float[]> DrawingElementsCoords { get; set; } = new Dictionary<string, float[]>();


        public BOM_Extractor()
        {
            //Задаём координаты инреесующих нас элементов в словарь DrawingElementsCoords
            //Сейчас захардкодено, но в дальнейшем можно передавать эту информацию откуда угодно: БД / конфиги
            DrawingElementsCoords.Add("Чертёж", new float[] { 869, 15, 224, 27 });
            DrawingElementsCoords.Add("Номер ревизии", new float[] { 1074, 85, 19, 14 });
            DrawingElementsCoords.Add("Лист", new float[] { 1034, 85, 24, 14 });
            DrawingElementsCoords.Add("Линия", new float[] { 1034, 99, 59, 13 });
            DrawingElementsCoords.Add("Ст. зона", new float[] { 985, 99, 49, 13 });

            //Добавляем столбцы в выходную таблицу
            BomDataTable.Columns.Add("№", typeof(int));
            BomDataTable.Columns.Add("Чертёж", typeof(string));
            BomDataTable.Columns.Add("Ст. зона", typeof(string));
            BomDataTable.Columns.Add("Линия", typeof(string));
            BomDataTable.Columns.Add("Лист", typeof(string));
            BomDataTable.Columns.Add("Номер ревизии", typeof(string));
            BomDataTable.Columns.Add("Категория компонента", typeof(string));
            BomDataTable.Columns.Add("Номенклатура", typeof(string));
            BomDataTable.Columns.Add("ММ", typeof(string));
            BomDataTable.Columns.Add("Идент. код", typeof(string));
            BomDataTable.Columns.Add("Количество", typeof(string));
            BomDataTable.Columns.Add("Ед. изм", typeof(string));

            BomDataTable.Columns["№"].AutoIncrement = true; //Столбец № имеет автоинкремент
            BomDataTable.Columns["№"].AutoIncrementSeed = 1; //Начинает с единицы
            BomDataTable.Columns["№"].AutoIncrementStep = 1; //Шаг инкремента 1
        }

        //Приватная функция, недоступная для вызова извне класса. Принимает в себя виртуальный PdfDocument и прямоугольник для извлечения текста
        private string GetTextFromArea(PdfDocument PdfDoc, Rectangle Rectan)
        {
            var Page = PageID == 0 ? PdfDoc.GetFirstPage() : PdfDoc.GetPage(PageID); //Берём номер страницы из параметра класса. Если этот параметр не задан, берём первую страницу
            var Filter = new IEventFilter[1]; //Задаём фильтр событий iText
            Filter[0] = new TextRegionEventFilter(Rectan); //Задаём текстовый фильтр событий для нашего прямоугольника
            var FilteredTextEventListener = new FilteredTextEventListener(new LocationTextExtractionStrategy(), Filter); //Задаём стратегию извлечения текста
            var Result = PdfTextExtractor.GetTextFromPage(Page, FilteredTextEventListener); //Извлекаем текст из прямоугольника
            return Result.Trim(); //Возвращаем тримленный извлечённый текст
        }

        //Функция, которая проверяет, является ли строка категорией компонента
        public bool IsComponentCategory(string s)
        {
            //Если строка whitespace, это не категория компонента
            if (string.IsNullOrWhiteSpace(s))
                return false;
            //Вернуть истину, если все символы строки либо буквы-апперкейсы либо вайтспейсы
            return s.ToCharArray().All(r => (char.IsUpper(r) && char.IsLetter(r)) || char.IsWhiteSpace(r));
        }

        public void ExtractRawData() //Функция, извлекающая текст в списки Components & Parameters
        {
            //Создаём витруальный PdfDocument, и с помощью PdfReader
            //загружаем в него содержимое .pdf, который находится по ссылке FilePath
            //он существует только в пределах соответствующей директивы using
            using (PdfDocument PdfDoc = new PdfDocument(new PdfReader(FilePath)))
            {
                //Чтобы задать прямоугольник для извлечения текста, требуется 4 параметра:
                //1. Координата X
                //2. Координата Y
                //3. Ширина
                //4. Высота

                //В цикле создаём прямоугольники, из которых считываем текст
                //Каждая строка расположена на 10 точек ниже предыдущий
                //Начиная с Y == 780 (начало области целевой таблицы) до Y == 270 (конец области целевой таблицы) считываем области с шагом -10 точек
                for (int i = 780; i >= 270; i -= 10)
                {
                    //Создаем прямоугольник для компонента по указанным координатам. Ось X, ширина и высота остаются одинаковыми, меняется только координата Y в соответствии с итератором i
                    Rectangle ComponentRectan = new Rectangle(833.5f, i, 180, 10);
                    //То же самое для параметров
                    Rectangle ParameterRectan = new Rectangle(1015, i, 159, 10);
                    //Передаем наш PdfDoc и прямоугольники в функцию по извлечению текста из области
                    string ComponentString = GetTextFromArea(PdfDoc, ComponentRectan);
                    string ParameterString = GetTextFromArea(PdfDoc, ParameterRectan);
                    //Добавляем полученные строки в списки компонентов и параметров соответственно
                    Components.Add(ComponentString);
                    Parameters.Add(ParameterString);
                }
            }
        }

        //Парсим выходную таблицу
        public void ParseBomDataTable()
        {
            //Открываем "виртуальный" PdfDoc
            using (PdfDocument PdfDoc = new PdfDocument(new PdfReader(FilePath)))
            {
                //Итерируем по каждой паре ключ-значение в словаре DrawingElementsCoords
                foreach (var kvp in DrawingElementsCoords)
                {
                    //Определяем прямоугольник по координатам значения словаря
                    Rectangle Rectan = new Rectangle(kvp.Value[0], kvp.Value[1], kvp.Value[2], kvp.Value[3]);
                    //Извлекаем текст из прямоугольника
                    string AreaText = GetTextFromArea(PdfDoc, Rectan);
                    //Задаем значение по-умоланию для столбца, который называется так же как извлекаемая область
                    BomDataTable.Columns[kvp.Key].DefaultValue = AreaText;
                    //Заполняем словарь DrawingElements полученным значением
                    DrawingElements[kvp.Key] = AreaText;
                }
            }

            //Выбрать индексы строк, которые являются началом описания компонента
            //Это те строки в Parameters, которые не пустые
            var NewComponentIndexes = Parameters.Select((v, i) => new { Index = i, Value = v })
                    .Where(r => !string.IsNullOrWhiteSpace(r.Value)).Select(r => r.Index).ToArray();

            //По каждому индексу
            foreach (var index in NewComponentIndexes)
            {
                //Вычисляем следующий индекс, чтобы понять, до какой строки продолжается описание текущего компонента
                //Если следующего индекса нет, значит это последний компонент, и нужно брать строки до конца списка
                var next_index = NewComponentIndexes.Where(r => r > index).Any() ? NewComponentIndexes.Where(r => r > index).First() : Parameters.Count - 1;
                //Среди списка строк Components выбрать индекс строки и значение строки...
                var ComponentDescriptionStrings = Components.Select((v, i) => new { Index = i, Value = v })
                    .Where(r => r.Index >= index & r.Index < next_index) //.. где индекс строки >= индекс строки начала описания текущего компонента И < индекса следующего компонента
                    .Where(r => !IsComponentCategory(r.Value)) // где строка не является категорией компонента
                    .Select(r => r.Value).ToArray(); //Выбрать значение этих строк в массив
                //Полученный массив склеить в строку
                var ComponentDescription = string.Join(" ", ComponentDescriptionStrings).Trim();
                //Получить категорию компонента
                //Это первая строка до начала описания текущего компонента, которая является категорий компонента
                var ComponentCategory = Components.Select((v, i) => new { Index = i, Value = v })
                    .Where(r => r.Index < index & IsComponentCategory(r.Value))
                    .Last().Value;

                //Создаём новый ряд нашей таблицы и добавляем все данные
                var row = BomDataTable.NewRow();
                row["Номенклатура"] = ComponentDescription;
                row["Категория компонента"] = ComponentCategory;
                var ParametersList = Parameters[index].Split(null).ToList();
                row["ММ"] = ParametersList[0].Trim();
                row["Идент. код"] = ParametersList[1].Trim();
                row["Количество"] = ParametersList[2].Trim();
                //Задаем Ед. изм согласно условиям ТЗ
                row["Ед. изм"] = ParametersList.Count == 4 ? ParametersList[3].Trim() : "PIECE";
                BomDataTable.Rows.Add(row);
            }
        }
    }
}