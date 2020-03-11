using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;


namespace GeneratorNaglowka
{
    public partial class Kontrolka : UserControl
    {
        private PowerPoint.Application taAplikacja;
        private PowerPoint.Presentation projekt;

        Color kolorKsztaltu = Color.LightGreen;
        Color kolorCzcionki = Color.White;
        Color kolorCzcionkiAlt = Color.Black;
        Color kolorKsztaltuAlt = Color.Pink;

        int przesuniecie = 0;

        public Kontrolka()
        {
            InitializeComponent();


            //czcionki
            Czcionka.DataSource = FontFamily.Families.Select(f => f.Name).ToList();
            Czcionka.SelectedIndex = 1;
            CzcionkaAlt.DataSource = FontFamily.Families.Select(f => f.Name).ToList();
            CzcionkaAlt.SelectedIndex = 1;
            Ksztalt.DataSource = Enum.GetValues(typeof(Office.MsoAutoShapeType));
            KsztaltAlt.DataSource = Enum.GetValues(typeof(Office.MsoAutoShapeType));

            // teksty
            for (int a = 0; a < SlajdNum.Maximum; a++)
            {
                KreatorBox("textBox", a, 0, 100, 20, GroupBox4, "Figura " + (a + 1));

                //znaczniki stron 'od'
                KreatorBox("odStrony", a, 120, 20, 20, GroupBox4, null);

                //znaczniki stron 'do'

                KreatorBox("doStrony", a, 160, 20, 20, GroupBox4, null);

            }

            GroupBox4.Controls["textBox" + 0].Show();
            GroupBox4.Controls["odStrony" + 0].Show();
            GroupBox4.Controls["doStrony" + 0].Show();


        }

        private void StartBut_Click(object sender, EventArgs e)
        {
            taAplikacja = Globals.ThisAddIn.Application;
            projekt = taAplikacja.ActivePresentation;
            PowerPoint.Slide Slajd = null;
            try //pobiera aktywny slajd, ustawia slajd na pierwszy w przypadku kursora pomiędzy slajdami
            {
                Slajd = taAplikacja.ActiveWindow.View.Slide;    
            }
            catch
            {
                Slajd = taAplikacja.ActivePresentation.Slides[1];
            }
            float slideWidth = projekt.PageSetup.SlideWidth;
            float slideHeight = projekt.PageSetup.SlideHeight; //pobiera szerokość, wysokość strony do dostosowania szerokości pól

            int war = SlajdNum.Value; //wartosc trackbara
            float marX = 10;
            float marY = 10; //marginesy
            if (MarginesX.Text != "") { float.TryParse(MarginesX.Text, out marX); }
            if (MarginesY.Text != "") { float.TryParse(MarginesY.Text, out marY); }

            float odstep = 0;
            if (Odstep.Text != "") { float.TryParse(Odstep.Text, out odstep); }

            float wys = 50;
            if (Wysokosc.Text != "") { float.TryParse(Wysokosc.Text, out wys); }
            float szer = (slideWidth - 2 * marX - war*odstep) / war; // atrybuty kształtu

            //granice działania
            int odSlajdu = 1, doSlajdu = projekt.Slides.Count;
            Int32.TryParse(Controls["slajdmin"].Text, out odSlajdu);
            if (odSlajdu == 0) odSlajdu = 1;
            Int32.TryParse(Controls["slajdmax"].Text, out doSlajdu);
            if (doSlajdu == 0) doSlajdu = projekt.Slides.Count;

            int slajdpoczatkowy = Slajd.SlideIndex;
            //pętla iterująca przez slajdy
            for (int slajd = odSlajdu; slajd <= doSlajdu; slajd++)
            {
                taAplikacja.ActiveWindow.View.GotoSlide(slajd);
                Slajd = taAplikacja.ActiveWindow.View.Slide;
                Int32.TryParse(GroupBox2.Controls["fontSize"].Text, out int fSize);
                if (fSize == 0) fSize = 26;
                //pętla tworząca figury
                for (int a = 0; a < war; a++)
                {

                    //resetowanie parametrow (potrzebne)
                    Office.MsoAutoShapeType ksztalt = (Office.MsoAutoShapeType)Ksztalt.SelectedValue;
                    string czcionka = Czcionka.SelectedValue.ToString();
                    Color tempKsztalt = kolorKsztaltu;
                    float gradient = 0.9f;
                    Color tempCzcionka = kolorCzcionki;
                    bool pogrubienie = false, kursywa = false;
                    //sprawdzanie zaznaczenia
                    Int32.TryParse(GroupBox4.Controls["odStrony" + a].Text, out int min);
                    Int32.TryParse(GroupBox4.Controls["doStrony" + a].Text, out int max);
                    if (Slajd.SlideIndex >= min && Slajd.SlideIndex <= max)
                    {
                        if (AltGrad.Checked) gradient = 0.5f;
                        if (AltKszt.Checked) ksztalt = (Office.MsoAutoShapeType)(KsztaltAlt.SelectedValue);
                        if (AltColorKszt.Checked) tempKsztalt = kolorKsztaltuAlt;
                        if (AltColorFont.Checked) tempCzcionka = kolorCzcionkiAlt;
                        if (AltFont.Checked) czcionka = CzcionkaAlt.SelectedValue.ToString();
                        if (Bold.Checked) pogrubienie = true;
                        if (Italic.Checked) kursywa = true;
                    }
                    //wywołanie funkcji tworzącej figury
                    KreatorKsztalt(Slajd, ksztalt, a, marX, marY, odstep, szer, wys, tempKsztalt, gradient, czcionka, fSize, tempCzcionka, pogrubienie, kursywa);
                }
            }
            taAplikacja.ActiveWindow.View.GotoSlide(slajdpoczatkowy);
        }

        private void DelBut_Click(object sender, EventArgs e)
        {
            //iteruje przez slajdy kasując obiekty zawierające nazwę "Generator_"
            taAplikacja = Globals.ThisAddIn.Application;
            //granice działania
            int odslajdu = 1, doslajdu = taAplikacja.ActivePresentation.Slides.Count;
            if (Controls["slajdmin"].Text != "") Int32.TryParse(Controls["slajdmin"].Text, out odslajdu);
            if (Controls["slajdmax"].Text != "") Int32.TryParse(Controls["slajdmax"].Text, out doslajdu);

            int slajdpoczatkowy = taAplikacja.ActiveWindow.View.Slide.SlideIndex;

            for (int slajd = odslajdu; slajd <= doslajdu; slajd++)
            {
                taAplikacja.ActiveWindow.View.GotoSlide(slajd);
                int obiekty = taAplikacja.ActiveWindow.View.Slide.Shapes.Count();
                while (obiekty > 0)
                {
                    PowerPoint.Shape kszt = taAplikacja.ActiveWindow.View.Slide.Shapes(obiekty);
                    if (kszt.Name.Contains("Generator_")) { kszt.Delete(); }
                    obiekty--;
                }
            }
            taAplikacja.ActiveWindow.View.GotoSlide(slajdpoczatkowy);
        }

        private void SlajdNum_Scroll(object sender, EventArgs e)
        {
            //wyświetla pola tekstowe i przesuwa granice grupy
            int wyswietlane = 0;
            while (wyswietlane < SlajdNum.Maximum)
            {
                Control temp = GroupBox4.Controls["textBox" + wyswietlane];
                if (wyswietlane < SlajdNum.Value) temp.Show();
                else temp.Hide();
                temp = GroupBox4.Controls["odStrony" + wyswietlane];
                if (wyswietlane < SlajdNum.Value) temp.Show();
                else temp.Hide();
                temp = GroupBox4.Controls["doStrony" + wyswietlane];
                if (wyswietlane < SlajdNum.Value) temp.Show();
                else temp.Hide();
                wyswietlane++;
            }
            GroupBox4.Size = new Size(GroupBox4.Width, GroupBox4.Controls["textBox" + (SlajdNum.Value - 1)].Location.Y + 30);
            for (int i = 3; i < 7; i++)
            {
                Grupa.Controls["GroupBox" + i].Location = new Point(6, Grupa.Controls["GroupBox" + (i - 1)].Location.Y + Grupa.Controls["GroupBox" + (i - 1)].Size.Height);
            }
        }

        //pola do zmiany koloru

        private Color ZmienKolor()
        {
            ColorDialog dialogchooser;
            dialogchooser = new ColorDialog();
            dialogchooser.AllowFullOpen = true;
            dialogchooser.ShowHelp = true;
            dialogchooser.ShowDialog();
            return dialogchooser.Color;
        }

        private void KolorCz_Click(object sender, EventArgs e)
        {
            kolorCzcionki=ZmienKolor();
        }

        private void KolorKszt_Click(object sender, EventArgs e)
        {
            kolorKsztaltu=ZmienKolor();
        }

        private void KolorKsztAlt_Click(object sender, EventArgs e)
        {
            kolorKsztaltuAlt=ZmienKolor();
        }

        private void KolorCzAlt_Click(object sender, EventArgs e)
        {
            kolorCzcionkiAlt=ZmienKolor();
        }

        //wyświetlanie dodatowych pól kiedy jest taka potrzeba

        private void AltColorKszt_CheckedChanged(object sender, EventArgs e)
        {
            KolorKsztAlt.Visible = !KolorKsztAlt.Visible;
        }

        private void AltFont_CheckedChanged(object sender, EventArgs e)
        {
            CzcionkaAlt.Visible = !CzcionkaAlt.Visible;
        }

        private void AltColorFont_CheckedChanged(object sender, EventArgs e)
        {
            KolorCzAlt.Visible = !KolorCzAlt.Visible;
        }

        private void AltKszt_CheckedChanged(object sender, EventArgs e)
        {
            KsztaltAlt.Visible = !KsztaltAlt.Visible;
        }

        //funkcje

        private PowerPoint.Shape KreatorKsztalt(PowerPoint.Slide Slajd, Office.MsoAutoShapeType ksztalt, int numer, float marginesX, float marginesY, float odstep, float szerokosc, float wysokosc, Color kolor, float gradient, string czcionka, float rozmiar, Color kolorCz, bool pogrubienie, bool kursywa)
        {
            //tworzy figure

            PowerPoint.Shape figura;
            figura = Slajd.Shapes.AddShape(ksztalt, marginesX + numer * (szerokosc + odstep) + odstep / 2, marginesY, szerokosc, wysokosc);
            figura.Name = "Generator_" + numer;
            //tekst
            Control temp = GroupBox4.Controls["textBox" + numer];
            figura.TextFrame.TextRange.Text = temp.Text;
            figura.TextFrame.TextRange.Font.Size = rozmiar;
            figura.TextFrame.TextRange.Font.Name = czcionka;
            figura.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(kolorCz.B, kolorCz.G, kolorCz.R).ToArgb();
            if (pogrubienie) figura.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            if (kursywa) figura.TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
            //figura
            figura.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1, 0);
            figura.Fill.ForeColor.RGB = Color.FromArgb(kolor.B, kolor.G, kolor.R).ToArgb();
            figura.Fill.BackColor.RGB = Color.FromArgb((int)(kolor.B * gradient), (int)(kolor.G * gradient), (int)(kolor.R * gradient)).ToArgb();

            return figura;
        }

        private TextBox KreatorBox(string nazwa, int a, int offset, int szerokosc, int wysokosc, GroupBox grupa, string tekst)
        {
            int X = SlajdNum.Location.X, Y = label3.Location.Y + 25;
            TextBox textBox = new TextBox();
            textBox.Location = new Point(X+offset, Y + a * wysokosc);
            textBox.Name = nazwa + a;
            textBox.Size = new Size(szerokosc, wysokosc);
            if(tekst!=null) textBox.Text = tekst;
            grupa.Controls.Add(textBox);
            textBox.BringToFront();
            textBox.Hide();
            return textBox;
        }

        private int FindLast(ControlCollection box)
        {
            //znajduje pozycje ostatniego elementu w grupie w celu ustalenia wysokosci
            int max = 0, temp = 0; ;
            for (int a = 0; a < box.Count; a++)
            {
                if (box[a].Visible)
                {
                    temp = box[a].Size.Height + box[a].Location.Y;
                    if (temp > max) max = temp;
                }
            }
            return max + 10;
        }

        private void Ogolny_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox check = (CheckBox)sender;
            if (check.Checked) check.Parent.Size = new Size(check.Parent.Size.Width, FindLast(check.Parent.Controls));
            else check.Parent.Size = new Size(check.Parent.Size.Width, 20);

            for (int i = 3; i < 7; i++)
            {
                Grupa.Controls["GroupBox" + i].Location = new Point (6,Grupa.Controls["GroupBox" + (i - 1)].Location.Y + Grupa.Controls["GroupBox" + (i - 1)].Size.Height);
            }

        }

        private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            int przesuniecieNowe = vScrollBar1.Value;
            for (int i = 2; i < 7; i++)
            {
                Grupa.Controls["GroupBox" + i].Location = new Point(6, Grupa.Controls["GroupBox" + (i)].Location.Y + przesuniecie*4 - przesuniecieNowe*4);        
            }
            przesuniecie = przesuniecieNowe;
        }
    }
}
