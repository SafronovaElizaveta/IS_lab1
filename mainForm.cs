using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace lab_1
{
    public partial class mainForm : Form
    {

        List<Car> DB = new List<Car>();

        string transmission = "", drive = "", brand = "", complect = "";
        long price;
        int year;
        Color color;
        List<int> criteria = new List<int>();
        public mainForm()
        {
            InitializeComponent();
            colorCB.DataSource = Color.colors;
            colorCB.DisplayMember = "Name";
            colorCB.ValueMember = "Id";

            DBLoader loader = new DBLoader();
            loader.load();
            DB = loader.GetDB();

            List<string> brands = new List<string> { "Любая" };
            foreach (var x in DB)
                if (brands.IndexOf(x.brand) == -1) brands.Add(x.brand);

            BrandCB.DataSource = brands;
            BrandCB.DisplayMember = "Name";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DBLoader loader = new DBLoader();
            loader.load();
            DB = loader.GetDB();
        }

        private void startBtn_Click(object sender, EventArgs e)
        {
            getDataFromForm();
            AutoClassResolver res = new AutoClassResolver(criteria);

            List<Car> ansList = DB.Where(x => (res.valideClasses().IndexOf(x.carClass) > -1) 
            && (brand.Equals("Любая") || x.brand.Equals(brand))
            && (complect.Equals("Любая") || x.complect.Equals(complect))
            && (x.price <= price) && (x.yearOfRelease >= year) 
            && (transmission.Equals("Any") || x.transmition.Equals(transmission))
            && (drive.Equals("Any") || x.drive.Equals(drive)) 
            && (color == Color.Any || x.color == color)).ToList();

            ansList.Sort(delegate (Car x, Car y) { return x.price.CompareTo(y.price); });

            string answer = ""; AnsTB.Text = "";
            foreach (var x in ansList)
                answer += x.color.Name + " " + x.brand + " " + x.nameCar + "\r\n" +
                    x.yearOfRelease + " года  " + x.drive + " " + x.transmition + "\r\n" +
                    x.price + " рублей \r\n\r\n";

            if (ansList.Count == 0)
                answer = "К сожалению, машина по вашему запросу не найдена";
            AnsTB.Text = answer;
        }

        private void getDataFromForm()
        {
            if (radioAT.Checked) transmission = "AT";
            if (radioMT.Checked) transmission = "MT";
            if (radioAnyKPP.Checked) transmission = "Any";
            if (radioAnyDrive.Checked) drive = "Any";
            if (radioFull.Checked) drive = "AWD";
            if (radioBack.Checked) drive = "RWD";
            if (radioFront.Checked) drive = "FWD";
            price = Convert.ToInt64(PriceTB.Text);
            year = Convert.ToInt32(YearTB.Text);
            color = (Color)colorCB.SelectedItem;
            brand = (string)BrandCB.SelectedItem;
            complect = (string)complectCB.SelectedItem;
            criteria.Clear();
            foreach (CheckBox x in criteriaPanel1.Controls)
                if (x.Checked)
                    criteria.Add(Convert.ToInt32(x.Tag));
        }
    }
}