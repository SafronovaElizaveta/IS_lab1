using System;
using System.Collections.Generic;

namespace lab_1
{
    class AutoClassResolver
    {
        private int checkedNumber = 0;
        List<int> checkedCriterion;
        Dictionary<string, double> autos = new Dictionary<string, double>();
        public AutoClassResolver(List<int> data)
        {
            checkedNumber = data.Count;
            checkedCriterion = data;
        }
        public List<string> valideClasses()
        {
            List<string> validClass = new List<string>();
            autos.Clear();
            autos.Add("Class_A", 0);
            autos.Add("Class_B", 0);
            autos.Add("Class_C", 0);
            autos.Add("Class_D", 0);
            autos.Add("Class_E", 0);
            autos.Add("Class_F", 0);
            autos.Add("Class_J", 0);  

            // критерий "Престижная"
            if (checkedCriterion.IndexOf(2) > -1)
            {
                autos["Class_D"] += 1.0 / checkedNumber;
                autos["Class_E"] += 1.0 / checkedNumber;
                autos["Class_F"] += 1.0 / checkedNumber;
            }

            // критерий "Спортивная"
            if (checkedCriterion.IndexOf(3) > -1)
            {
                autos["Class_F"] += 1.0 / checkedNumber;
                autos["Class_J"] += 1.0 / checkedNumber;
            }
            //  критерий "Компактная"
            if (checkedCriterion.IndexOf(1) > -1)
            {
                autos["Class_A"] += 1.0 / checkedNumber;
                autos["Class_C"] += 1.0 / checkedNumber;
                autos["Class_E"] += 1.0 / checkedNumber;
            }
            
            // критерий "Экономичная"
            if (checkedCriterion.IndexOf(5) > -1)
            {
                autos["Class_A"] += 1.0 / checkedNumber;
                autos["Class_B"] += 1.0 / checkedNumber;
                autos["Class_C"] += 1.0 / checkedNumber;
            }

            // критерий "Для отдыха"
            if (checkedCriterion.IndexOf(4) > -1)
            {
                autos["Class_B"] += 1.0 / checkedNumber;
                autos["Class_C"] += 1.0 / checkedNumber;
                autos["Class_E"] += 1.0 / checkedNumber;
            }

            foreach (var x in autos)
                if (checkedNumber == 0 || x.Value == Convert.ToInt32(1)) validClass.Add(x.Key);
            return validClass;
        }
    }

    class Car
    {
        public string brand;
        public string nameCar;
        public long price;
        public int yearOfRelease;
        public string transmition; // КПП
        public string drive; // привод
        public Color color;
        public string complect; //комплектация
        public string carClass;

        public Car(string brand, string name, long price, int year, string transmition, string drive, Color color, string complect, string carClass)
        {
            this.brand = brand;
            nameCar = name;
            this.price = price;
            yearOfRelease = year;
            this.transmition = transmition;
            this.drive = drive;
            this.color = color;
            this.complect = complect;
            this.carClass = carClass;
        }
    }

    class Color
    {
        public int Id { get; set; }
        public string Name { get; set; }

        public Color(int id, string name)
        {
            Id = id;
            Name = name;
        }

        public static readonly Color Any = new Color(1, "Любой");
        public static readonly Color White = new Color(2, "Белый");
        public static readonly Color Black = new Color(3, "Черный");
        public static readonly Color Red = new Color(4, "Красный");
        public static readonly Color Blue = new Color(5, "Синий");
        public static readonly Color Green = new Color(6, "Зеленый");
        public static readonly Color Grey = new Color(7, "Серый");
        public static readonly Color Yellow = new Color(8, "Желтый");
        public static readonly Color Orange = new Color(9, "Оранжевый");
        public static readonly List<Color> colors = new List<Color>
            {Any, White, Black, Red, Blue, Green, Grey,Yellow, Orange};

    }

    class DBLoader 
    {
        Sheet log = Sheet.getInstance();
        List<Car> DB = new List<Car>();


        // подгружает записи из excel
        public void load()
        {
              MyExcel worksheet = log.table;
              worksheet.GetRecord(DB);
        }
        public List<Car> GetDB()
        {
            return DB;
        }
    }

    class Sheet
    {
        private static Sheet instance;

        public MyExcel table = new MyExcel();

        protected Sheet()
        {
            this.table.ReadDoc();
        }

        public static Sheet getInstance()
        {
            if (instance == null)
                instance = new Sheet();
            return instance;
        }
    }

}
