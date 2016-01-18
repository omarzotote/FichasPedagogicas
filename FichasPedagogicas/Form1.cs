using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace FichasPedagogicas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Routing document
            string fileName = firstNameTB.Text + " " + lastNameTB.Text + " " + NameTB.Text + ".docx";
            string targetFolder = @"C:\Users\Omar\documents\visual studio 2015\Projects\FichasPedagogicas\FichasPedagogicas\Docs\";
            string copyFile = "file.docx";
            string sourcePath = @"c:\users\omar\documents\visual studio 2015\Projects\FichasPedagogicas\FichasPedagogicas\";
            string sourceFolder = System.IO.Path.Combine(sourcePath, copyFile);
            string destFolder = System.IO.Path.Combine(targetFolder, fileName);
            if (!System.IO.Directory.Exists(targetFolder))
            {
                System.IO.Directory.CreateDirectory(targetFolder);
            }
            System.IO.File.Copy(sourceFolder, destFolder, true);

            //refs
            object name = "name";
            object firstname = "firstname";
            object lastname = "lastname";
            object age = "age";
            object day = "day";
            object month = "month";
            object year = "year";
            object agecom = "agecom";
            object male = "male";
            object female = "female";
            object address = "address";
            object colony = "colony";
            object cp = "cp";
            object telephone = "telephone";
            object curp = "curp";
            object date = "date";
            object fname = "fname";
            object fjob = "fjob";
            object fjobaddress = "fjobaddress";
            object fphone = "fphone";
            object mname = "mname";
            object mjob = "mjob";
            object mjobaddress = "mjobaddress";
            object mphone = "mphone";
            object bno = "bno";
            object byes = "byes";
            object bgroup = "bgroup";
            object fe = "pe";
            object me = "me";
            object ft = "pt";
            object mt = "mt";
            object fn = "pn";
            object mn = "mn";
            object fm = "fm";
            object f = "f";
            object m = "m";
            object o = "o";
            object house = "house";
            object appartment = "appartment";
            object room = "room";
            object move = "move";
            object water = "water";
            object telephonenumber = "telephonenumber";
            object light = "light";
            object sewer = "sewer";
            object gas = "gas";
            object pavement = "pavement";
            object familysize = "familysize";
            object peopleworking = "peopleworking";
            object profit = "profit";
            object ownhouseyes = "ownhouseyes";
            object ownhouseno = "ownhouseno";
            object imss = "imss";
            object isste = "isste";
            object isea = "isea";
            object particular = "particular";
            object none = "none";
            object popular = "popular";
            object vaccineyes = "vaccineyes";
            object vaccineno = "vaccineno";
            object diseases = "diseases";
            object allergy = "allergy";
            object medicines = "medicines";
            object bloodtype = "bloodtype";
            object avn = "avn";
            object avd = "avd";
            object ave = "ave";
            object avp = "avp";
            object aan = "aan";
            object aad = "aad";
            object aae = "aae";
            object aap = "aap";
            object ppn = "ppn";
            object ppd = "ppd";
            object ppe = "ppe";
            object ppp = "ppp";
            object sports = "sports";
            object artistics = "artistics";
            object cultural = "cultural";
            object recreative = "recreative";
            object other = "other";
            object dad = "dad";
            object mom = "mom";
            object brothers = "brothers";
            object partners = "partners";
            object teachers = "teachers";
            object neighbors = "neighbors";


            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Open(destFolder);

            //get Bookmarks
            Word.Range name1 = oDoc.Bookmarks.get_Item(ref name).Range;
            Word.Range firstname1 = oDoc.Bookmarks.get_Item(ref firstname).Range;
            Word.Range lastname1 = oDoc.Bookmarks.get_Item(ref lastname).Range;
            Word.Range age1 = oDoc.Bookmarks.get_Item(ref age).Range;
            Word.Range day1 = oDoc.Bookmarks.get_Item(ref day).Range;
            Word.Range month1 = oDoc.Bookmarks.get_Item(ref month).Range;
            Word.Range year1 = oDoc.Bookmarks.get_Item(ref year).Range;
            Word.Range agecom1 = oDoc.Bookmarks.get_Item(ref agecom).Range;
            Word.Range male1 = oDoc.Bookmarks.get_Item(ref male).Range;
            Word.Range female1 = oDoc.Bookmarks.get_Item(ref female).Range;
            Word.Range address1 = oDoc.Bookmarks.get_Item(ref address).Range;
            Word.Range colony1 = oDoc.Bookmarks.get_Item(ref colony).Range;
            Word.Range cp1 = oDoc.Bookmarks.get_Item(ref cp).Range;
            Word.Range telephone1 = oDoc.Bookmarks.get_Item(ref telephone).Range;
            Word.Range curp1 = oDoc.Bookmarks.get_Item(ref curp).Range;
            Word.Range date1 = oDoc.Bookmarks.get_Item(ref date).Range;
            Word.Range fname1 = oDoc.Bookmarks.get_Item(ref fname).Range;
            Word.Range fjob1 = oDoc.Bookmarks.get_Item(ref fjob).Range;
            Word.Range fjobaddress1 = oDoc.Bookmarks.get_Item(ref fjobaddress).Range;
            Word.Range fphone1 = oDoc.Bookmarks.get_Item(ref fphone).Range;
            Word.Range mname1 = oDoc.Bookmarks.get_Item(ref mname).Range;
            Word.Range mjob1 = oDoc.Bookmarks.get_Item(ref mjob).Range;
            Word.Range mjobaddress1 = oDoc.Bookmarks.get_Item(ref mjobaddress).Range;
            Word.Range mphone1 = oDoc.Bookmarks.get_Item(ref mphone).Range;
            Word.Range bno1 = oDoc.Bookmarks.get_Item(ref bno).Range;
            Word.Range byes1 = oDoc.Bookmarks.get_Item(ref byes).Range;
            Word.Range bgroup1 = oDoc.Bookmarks.get_Item(ref bgroup).Range;
            Word.Range fe1 = oDoc.Bookmarks.get_Item(ref fe).Range;
            Word.Range me1 = oDoc.Bookmarks.get_Item(ref me).Range;
            Word.Range ft1 = oDoc.Bookmarks.get_Item(ref ft).Range;
            Word.Range mt1 = oDoc.Bookmarks.get_Item(ref mt).Range;
            Word.Range fn1 = oDoc.Bookmarks.get_Item(ref fn).Range;
            Word.Range mn1 = oDoc.Bookmarks.get_Item(ref mn).Range;
            Word.Range fm1 = oDoc.Bookmarks.get_Item(ref fm).Range;
            Word.Range f1 = oDoc.Bookmarks.get_Item(ref f).Range;
            Word.Range m1 = oDoc.Bookmarks.get_Item(ref m).Range;
            Word.Range o1 = oDoc.Bookmarks.get_Item(ref o).Range;
            Word.Range house1 = oDoc.Bookmarks.get_Item(ref house).Range;
            Word.Range appartment1 = oDoc.Bookmarks.get_Item(ref appartment).Range;
            Word.Range room1 = oDoc.Bookmarks.get_Item(ref room).Range;
            Word.Range move1 = oDoc.Bookmarks.get_Item(ref move).Range;
            Word.Range water1 = oDoc.Bookmarks.get_Item(ref water).Range;
            Word.Range telephonenumber1 = oDoc.Bookmarks.get_Item(ref telephonenumber).Range;
            Word.Range light1 = oDoc.Bookmarks.get_Item(ref light).Range;
            Word.Range sewer1 = oDoc.Bookmarks.get_Item(ref sewer).Range;
            Word.Range gas1 = oDoc.Bookmarks.get_Item(ref gas).Range;
            Word.Range pavement1 = oDoc.Bookmarks.get_Item(ref pavement).Range;
            Word.Range familysize1 = oDoc.Bookmarks.get_Item(ref familysize).Range;
            Word.Range peopleworking1 = oDoc.Bookmarks.get_Item(ref peopleworking).Range;
            Word.Range profit1 = oDoc.Bookmarks.get_Item(ref profit).Range;
            Word.Range ownhouseyes1 = oDoc.Bookmarks.get_Item(ref ownhouseyes).Range;
            Word.Range ownhouseno1 = oDoc.Bookmarks.get_Item(ref ownhouseno).Range;
            Word.Range imss1 = oDoc.Bookmarks.get_Item(ref imss).Range;
            Word.Range isste1 = oDoc.Bookmarks.get_Item(ref isste).Range;
            Word.Range isea1 = oDoc.Bookmarks.get_Item(ref isea).Range;
            Word.Range particular1 = oDoc.Bookmarks.get_Item(ref particular).Range;
            Word.Range none1 = oDoc.Bookmarks.get_Item(ref none).Range;
            Word.Range popular1 = oDoc.Bookmarks.get_Item(ref popular).Range;
            Word.Range vaccineyes1 = oDoc.Bookmarks.get_Item(ref vaccineyes).Range;
            Word.Range vaccineno1 = oDoc.Bookmarks.get_Item(ref vaccineno).Range;
            Word.Range diseases1 = oDoc.Bookmarks.get_Item(ref diseases).Range;
            Word.Range allergy1 = oDoc.Bookmarks.get_Item(ref allergy).Range;
            Word.Range medicines1 = oDoc.Bookmarks.get_Item(ref medicines).Range;
            Word.Range bloodtype1 = oDoc.Bookmarks.get_Item(ref bloodtype).Range;
            Word.Range avn1 = oDoc.Bookmarks.get_Item(ref avn).Range;
            Word.Range avd1 = oDoc.Bookmarks.get_Item(ref avd).Range;
            Word.Range ave1 = oDoc.Bookmarks.get_Item(ref ave).Range;
            Word.Range avp1 = oDoc.Bookmarks.get_Item(ref avp).Range;
            Word.Range aan1 = oDoc.Bookmarks.get_Item(ref aan).Range;
            Word.Range aad1 = oDoc.Bookmarks.get_Item(ref aad).Range;
            Word.Range aae1 = oDoc.Bookmarks.get_Item(ref aae).Range;
            Word.Range aap1 = oDoc.Bookmarks.get_Item(ref aap).Range;
            Word.Range ppn1 = oDoc.Bookmarks.get_Item(ref ppn).Range;
            Word.Range ppd1 = oDoc.Bookmarks.get_Item(ref ppd).Range;
            Word.Range ppe1 = oDoc.Bookmarks.get_Item(ref ppe).Range;
            Word.Range ppp1 = oDoc.Bookmarks.get_Item(ref ppp).Range;
            Word.Range sports1 = oDoc.Bookmarks.get_Item(ref sports).Range;
            Word.Range artistics1 = oDoc.Bookmarks.get_Item(ref artistics).Range;
            Word.Range cultural1 = oDoc.Bookmarks.get_Item(ref cultural).Range;
            Word.Range recreative1 = oDoc.Bookmarks.get_Item(ref recreative).Range;
            Word.Range other1 = oDoc.Bookmarks.get_Item(ref other).Range;
            Word.Range dad1 = oDoc.Bookmarks.get_Item(ref dad).Range;
            Word.Range mom1 = oDoc.Bookmarks.get_Item(ref mom).Range;
            Word.Range brothers1 = oDoc.Bookmarks.get_Item(ref brothers).Range;
            Word.Range partners1 = oDoc.Bookmarks.get_Item(ref partners).Range;
            Word.Range teachers1 = oDoc.Bookmarks.get_Item(ref teachers).Range;
            Word.Range neighbors1 = oDoc.Bookmarks.get_Item(ref neighbors).Range;

            //get textBox Text
            name1.Text = NameTB.Text;
            firstname1.Text = firstNameTB.Text;
            lastname1.Text = lastNameTB.Text;
            age1.Text = "  " + AgeTB.Text;
            day1.Text = "   " + DayTB.Text;
            month1.Text = "   " + MonthTB.Text;
            year1.Text = "   " + YearTB.Text;
            agecom1.Text = "       " + AgeComTB.Text;
            address1.Text = AddressTB.Text;
            colony1.Text = ColonyTB.Text;
            cp1.Text = CpTB.Text;
            telephone1.Text = TelephoneTB.Text;
            curp1.Text = CurpTB.Text;
            date1.Text = DateTB.Text;
            fname1.Text = FnameTB.Text;
            fjob1.Text = FjobTB.Text;
            fjobaddress1.Text = FjobAddressTB.Text;
            fphone1.Text = FPhoneTB.Text;
            mname1.Text = MnameTB.Text;
            mjob1.Text = MjobTB.Text;
            mjobaddress1.Text = MjobAddressTB.Text;
            mphone1.Text = MPhoneTB.Text;
            bgroup1.Text = BgroupTB.Text;
            if(felist.Text != "" && felist.Text != "N/A")
                fe1.Text = felist.Text.First().ToString();
            if (melist.Text != "" && melist.Text != "N/A")
                me1.Text = melist.Text.First().ToString();
            if (ftlist.Text != "" && ftlist.Text != "N/A")
                ft1.Text = ftlist.Text.First().ToString();
            if (mtlist.Text != "" && mtlist.Text != "N/A")
                mt1.Text = mtlist.Text.First().ToString();
            if (fnlist.Text != "" && fnlist.Text != "N/A")
                fn1.Text = fnlist.Text.First().ToString();
            if (mnlist.Text != "" && mnlist.Text != "N/A")
                mn1.Text = mnlist.Text.First().ToString();
            familysize1.Text = familyTB.Text;
            peopleworking1.Text = peopleTB.Text;
            profit1.Text = profitTB.Text;
            diseases1.Text = diseasesTB.Text;
            allergy1.Text = allergyTB.Text;
            medicines1.Text = medicinesTB.Text;
            bloodtype1.Text = bloodtypeTB.Text;
            sports1.Text = sportsTB.Text;
            artistics1.Text = artisticsTB.Text;
            cultural1.Text = culturalTB.Text;
            recreative1.Text = recreativeTB.Text;
            other1.Text = otherTB.Text;

            //Bad boy :(
            if (!dadCB.Checked)
                dad1.ShapeRange[1].Delete();
            if (!momCB.Checked)
                mom1.ShapeRange[1].Delete();
            if (!brothersCB.Checked)
                brothers1.ShapeRange[1].Delete();
            if (!partnersCB.Checked)
                partners1.ShapeRange[1].Delete();
            if (!teachersCB.Checked)
                teachers1.ShapeRange[1].Delete();
            if (!neighborsCB.Checked)
                neighbors1.ShapeRange[1].Delete();


            //House Type
            string typeHouse = houseType.Text;
            switch (typeHouse)
            {
                case "Casa":
                    appartment1.ShapeRange[1].Delete();
                    room1.ShapeRange[1].Delete();
                    move1.ShapeRange[1].Delete();
                    break;
                case "Departamento o vecindad":
                    house1.ShapeRange[1].Delete();
                    room1.ShapeRange[1].Delete();
                    move1.ShapeRange[1].Delete();
                    break;
                case "Cuarto":
                    appartment1.ShapeRange[1].Delete();
                    house1.ShapeRange[1].Delete();
                    move1.ShapeRange[1].Delete();
                    break;
                case "Vivienda móvil":
                    appartment1.ShapeRange[1].Delete();
                    room1.ShapeRange[1].Delete();
                    house1.ShapeRange[1].Delete();
                    break;
                default:
                    appartment1.ShapeRange[1].Delete();
                    room1.ShapeRange[1].Delete();
                    house1.ShapeRange[1].Delete();
                    move1.ShapeRange[1].Delete();
                    break;
            }

            //Disabilities
            string av = avList.Text;
            switch (av)
            {
                case "Normal":
                    avd1.ShapeRange[1].Delete();
                    ave1.ShapeRange[1].Delete();
                    avp1.ShapeRange[1].Delete();
                    break;
                case "Detectado":
                    avn1.ShapeRange[1].Delete();
                    ave1.ShapeRange[1].Delete();
                    avp1.ShapeRange[1].Delete();
                    break;
                case "En tratamiento":
                    avn1.ShapeRange[1].Delete();
                    avd1.ShapeRange[1].Delete();
                    avp1.ShapeRange[1].Delete();
                    break;
                case "Problema Resuelto":
                    avd1.ShapeRange[1].Delete();
                    ave1.ShapeRange[1].Delete();
                    avn1.ShapeRange[1].Delete();
                    break;
                default:
                    avd1.ShapeRange[1].Delete();
                    ave1.ShapeRange[1].Delete();
                    avn1.ShapeRange[1].Delete();
                    avp1.ShapeRange[1].Delete();
                    break;
            }
            string aa = aaList.Text;
            switch (aa)
            {
                case "Normal":
                    aad1.ShapeRange[1].Delete();
                    aae1.ShapeRange[1].Delete();
                    aap1.ShapeRange[1].Delete();
                    break;
                case "Detectado":
                    aan1.ShapeRange[1].Delete();
                    aae1.ShapeRange[1].Delete();
                    aap1.ShapeRange[1].Delete();
                    break;
                case "En tratamiento":
                    aan1.ShapeRange[1].Delete();
                    aad1.ShapeRange[1].Delete();
                    aap1.ShapeRange[1].Delete();
                    break;
                case "Problema Resuelto":
                    aad1.ShapeRange[1].Delete();
                    aae1.ShapeRange[1].Delete();
                    aan1.ShapeRange[1].Delete();
                    break;
                default:
                    aad1.ShapeRange[1].Delete();
                    aae1.ShapeRange[1].Delete();
                    aan1.ShapeRange[1].Delete();
                    aap1.ShapeRange[1].Delete();
                    break;
            }
            string pp = ppList.Text;
            switch (pp)
            {
                case "Normal":
                    ppd1.ShapeRange[1].Delete();
                    ppe1.ShapeRange[1].Delete();
                    ppp1.ShapeRange[1].Delete();
                    break;
                case "Detectado":
                    ppn1.ShapeRange[1].Delete();
                    ppe1.ShapeRange[1].Delete();
                    ppp1.ShapeRange[1].Delete();
                    break;
                case "En tratamiento":
                    ppn1.ShapeRange[1].Delete();
                    ppd1.ShapeRange[1].Delete();
                    ppp1.ShapeRange[1].Delete();
                    break;
                case "Problema Resuelto":
                    ppd1.ShapeRange[1].Delete();
                    ppe1.ShapeRange[1].Delete();
                    ppn1.ShapeRange[1].Delete();
                    break;
                default:
                    ppd1.ShapeRange[1].Delete();
                    ppe1.ShapeRange[1].Delete();
                    ppn1.ShapeRange[1].Delete();
                    ppp1.ShapeRange[1].Delete();
                    break;
            }
            //Vaccines
            string vaccines = vaccinesList.Text;
            switch(vaccines)
            {
                case "Sí":
                    vaccineno1.ShapeRange[1].Delete();
                    break;
                case "No":
                    vaccineyes1.ShapeRange[1].Delete();
                    break;
                default:
                    vaccineyes1.ShapeRange[1].Delete();
                    vaccineno1.ShapeRange[1].Delete();
                    break;
            }

            //Medic Service
            if (!imssCB.Checked)
                imss1.ShapeRange[1].Delete();
            if (!issteCB.Checked)
                isste1.ShapeRange[1].Delete();
            if (!iseaCB.Checked)
                isea1.ShapeRange[1].Delete();
            if (!particularCB.Checked)
                particular1.ShapeRange[1].Delete();
            if (!noneCB.Checked)
                none1.ShapeRange[1].Delete();
            if (!popularCB.Checked)
                popular1.ShapeRange[1].Delete();

            //Own House
            string ownHouseStr = ownHouse.Text;
            switch (ownHouseStr)
            {
                case "Sí":
                    ownhouseno1.ShapeRange[1].Delete();
                    break;
                case "No":
                    ownhouseyes1.ShapeRange[1].Delete();
                    break;
                default:
                    ownhouseno1.ShapeRange[1].Delete();
                    ownhouseyes1.ShapeRange[1].Delete();
                    break;
            }

            //House Services
            if(!waterCB.Checked)
                water1.ShapeRange[1].Delete();
            if(!phoneCB.Checked)
                telephonenumber1.ShapeRange[1].Delete();
            if(!lightCB.Checked)
                light1.ShapeRange[1].Delete();
            if (!sewerCB.Checked)
                sewer1.ShapeRange[1].Delete();
            if (!gasCB.Checked)
                gas1.ShapeRange[1].Delete();
            if (!pavementCB.Checked)
                pavement1.ShapeRange[1].Delete();


            //Live with
            string family = livesWith.Text;
            switch(family)
            {
                case "Padre y madre":
                    f1.ShapeRange[1].Delete();
                    m1.ShapeRange[1].Delete();
                    o1.ShapeRange[1].Delete();
                    break;
                case "Padre":
                    fm1.ShapeRange[1].Delete();
                    m1.ShapeRange[1].Delete();
                    o1.ShapeRange[1].Delete();
                    break;
                case "Madre":
                    fm1.ShapeRange[1].Delete();
                    f1.ShapeRange[1].Delete();
                    o1.ShapeRange[1].Delete();
                    break;
                case "Otro":
                    fm1.ShapeRange[1].Delete();
                    f1.ShapeRange[1].Delete();
                    m1.ShapeRange[1].Delete();
                    break;
                default:
                    fm1.ShapeRange[1].Delete();
                    f1.ShapeRange[1].Delete();
                    m1.ShapeRange[1].Delete();
                    o1.ShapeRange[1].Delete();
                    break;
            }

            //Male-Female
            string gender = genderList.Text;
            switch (gender)
            {
                case "Hombre":
                    female1.ShapeRange[1].Delete();
                    break;
                case "Mujer":
                    male1.ShapeRange[1].Delete();
                    break;
                default:
                    female1.ShapeRange[1].Delete();
                    male1.ShapeRange[1].Delete();
                    break;
            }

            //Brothers
            string hasBrothers = brothersList.Text;
            switch(hasBrothers)
            {
                case "Sí":
                    bno1.ShapeRange[1].Delete();
                    break;
                case "No":
                    byes1.ShapeRange[1].Delete();
                    break;
                default:
                    byes1.ShapeRange[1].Delete();
                    bno1.ShapeRange[1].Delete();
                    break;
            }

            //Create objects
            object wname = name1;
            object wfirstname = firstname1;
            object wlastname = lastname1;
            object wage = age1;
            object wday = day1;
            object wmonth = month1;
            object wyear = year1;
            object wagecom = agecom1;
            object waddress = address1;
            object wcolony = colony1;
            object wcp = cp1;
            object wtelephone = telephone1;
            object wcurp = curp1;
            object wdate = date1;
            object wfname = fname1;
            object wfjob = fjob1;
            object wfjobaddress = fjobaddress1;
            object wfphone = fphone1;
            object wmname = mname1;
            object wmjob = mjob1;
            object wmjobaddress = mjobaddress1;
            object wmphone = mphone1;
            object wbgroup = bgroup1;
            object wfe = fe1;
            object wme = me1;
            object wft = ft1;
            object wmt = mt1;
            object wfn = fn1;
            object wmn = mn1;
            object wdiseases = diseases1;
            object wallergy = allergy1;
            object wmedicines = medicines1;
            object wbloodtype = bloodtype1;
            object wsports = sports1;
            object wartistics = artistics1;
            object wcultural = cultural1;
            object wrecreative = recreative1;
            object wother = other1;
            
            //Adding bookmarks
            oDoc.Bookmarks.Add("name", ref wname);
            oDoc.Bookmarks.Add("firstname", ref wfirstname);
            oDoc.Bookmarks.Add("lastname", ref wlastname);
            oDoc.Bookmarks.Add("age", ref wage);
            oDoc.Bookmarks.Add("day", ref wday);
            oDoc.Bookmarks.Add("month", ref wmonth);
            oDoc.Bookmarks.Add("year", ref wyear);
            oDoc.Bookmarks.Add("agecom", ref wagecom);
            oDoc.Bookmarks.Add("address", ref waddress);
            oDoc.Bookmarks.Add("colony", ref wcolony);
            oDoc.Bookmarks.Add("cp", ref wcp);
            oDoc.Bookmarks.Add("telephone", ref wtelephone);
            oDoc.Bookmarks.Add("curp", ref wcurp);
            oDoc.Bookmarks.Add("date", ref wdate);
            oDoc.Bookmarks.Add("fname", ref wfname);
            oDoc.Bookmarks.Add("fjob", ref wfjob);
            oDoc.Bookmarks.Add("fjobaddress", ref wfjobaddress);
            oDoc.Bookmarks.Add("fphone", ref wfphone);
            oDoc.Bookmarks.Add("mname", ref wmname);
            oDoc.Bookmarks.Add("mjob", ref wmjob);
            oDoc.Bookmarks.Add("mjobaddress", ref wmjobaddress);
            oDoc.Bookmarks.Add("mphone", ref wmphone);
            oDoc.Bookmarks.Add("bgroup", ref wbgroup);
            oDoc.Bookmarks.Add("fe", ref wfe);
            oDoc.Bookmarks.Add("me", ref wme);
            oDoc.Bookmarks.Add("ft", ref wft);
            oDoc.Bookmarks.Add("mt", ref wmt);
            oDoc.Bookmarks.Add("fn", ref wfn);
            oDoc.Bookmarks.Add("mn", ref wmn);
            oDoc.Bookmarks.Add("diseases", ref wdiseases);
            oDoc.Bookmarks.Add("allergy", ref wallergy);
            oDoc.Bookmarks.Add("medicines", ref wmedicines);
            oDoc.Bookmarks.Add("bloodtype", ref wbloodtype);
            oDoc.Bookmarks.Add("sports", ref wsports);
            oDoc.Bookmarks.Add("artistics", ref wartistics);
            oDoc.Bookmarks.Add("cultural", ref wcultural);
            oDoc.Bookmarks.Add("recreative", ref wrecreative);
            oDoc.Bookmarks.Add("other", ref wother);


            //Oppening word document
            oWord.Visible = true;

            //Clear Textbox
            foreach (Control X in this.Controls)
            {
                if (X is TextBox)
                    X.Text = "";
            }
        }

        private void aaList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void brothersCB_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
    }
}
