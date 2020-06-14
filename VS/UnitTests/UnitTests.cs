using Microsoft.VisualStudio.TestTools.UnitTesting;
using NIPIGAS_BomExtractorLogic;

namespace UnitTests
{
    [TestClass] //Класс для юнит-тестов
    public class UnitTests
    {
        private const string MyPDF = @"..\..\MyPDF\PipeDrawing.pdf"; //относительный путь к тестовому PDF

        [TestMethod] //Тестовый метод
        public void MainTest()
        {
            var BOME = new BOM_Extractor(); //Инициализиуем новый инстанс класса с основной логикой
            BOME.FilePath = MyPDF; //Передаём в экземпляр класса путь к тестовому PDF
            //BOME.PageID = 1; //Передаём в экземпляр класса номер целевой страницы
            BOME.ExtractRawData(); //Запускаем процесс извлечения компонентов и параметров из целевой таблицы
            BOME.ParseBomDataTable(); //Парсим таблицу
            var BomDT = BOME.BomDataTable;
            var DrawingElements = BOME.DrawingElements;
        }
        [TestMethod] //Тестовый метод для функции IsComponentCategory
        public void CheckForComponentCategory()
        {
            var BOME = new BOM_Extractor();
            string Test1 = "PIPE SUPPORTS";
            string Test2 = "NREQD 125 - 250 AARH F10CFC07B005.";
            Assert.IsTrue(BOME.IsComponentCategory(Test1));
            Assert.IsFalse(BOME.IsComponentCategory(Test2));
        }
    }
}
