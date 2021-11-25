using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelComparison
{
    class Program
    {
        public static void CompareExcel(string path1, string path2, string rows1, string column1,string rows2,string column2, string sheet, bool isRUReport)
        {
            Console.WriteLine("\nLoading Sheets...\n");

            var workbook1 = WorkBook.Load(path1);
            var workbook2 = WorkBook.Load(path2);

            var worksheet1 = workbook1.GetWorkSheet(sheet); //mention name of sheet e.g. Sheet1
            var worksheet2 = workbook2.GetWorkSheet(sheet);

            Console.WriteLine("Sheets Loaded.");

            var allRows1 = new List<Row>();
            var allRows2 = new List<Row>();
            Row row;
            bool flag = false;
            int temp = 0;

            //Loading Excel to save output
            WorkBook workBook = WorkBook.Load("Comparison3.xlsx");
            var newSheet = workBook.GetWorkSheet("Articles");
            int rowNum = 2;

            //to count and check if it is entirely new record
            int isNew = 0;
            int entirelyNewRecord = 0;
            //int numOfRowsInSheet2FoundInSheet1 = 0;

            //uncomment no. of entries equal to no. of columns in excel
            for (int i=2; i <= Convert.ToInt32(rows1); i++)
            {
                var range1 = worksheet1[$"A{i}:{column1}{i}"].ToList();                
                row = new Row() {
                    Entry1 = range1[0].ToString(),
                    Entry2 = range1[1].ToString(),
                    Entry3 = range1[2].ToString(),
                    Entry4 = range1[3].ToString(),
                    Entry5 = range1[4].ToString(),
                    Entry6 = range1[5].ToString(),
                    Entry7 = range1[6].ToString(),
                    Entry8 = range1[7].ToString(),
                    Entry9 = range1[8].ToString(),
                    Entry10 = range1[9].ToString(),
                    Entry11 = range1[10].ToString(),
                    Entry12 = range1[11].ToString(),
                    Entry13 = range1[12].ToString(),
                    Entry14 = range1[13].ToString(),
                    Entry15 = range1[14].ToString(),
                    Entry16 = range1[15].ToString(),
                    Entry17 = range1[16].ToString(),
                    Entry18 = range1[17].ToString(),
                    Entry19 = range1[18].ToString(),
                    Entry20 = range1[19].ToString(),
                    Entry21 = range1[20].ToString(),
                    Entry22 = range1[21].ToString(),
                    Entry23 = range1[22].ToString(),
                    Entry24 = range1[23].ToString(),
                    Entry25 = range1[24].ToString(),
                    Entry26 = range1[25].ToString(),
                    Entry27 = range1[26].ToString(),
                    Entry28 = range1[27].ToString(),
                    Entry29 = range1[28].ToString(),
                    Entry30 = range1[29].ToString()
                };
                allRows1.Add(row);
            }

            for (int i = 2; i <= Convert.ToInt32(rows2); i++)
            {
                var range2 = worksheet2[$"A{i}:{column2}{i}"].ToList();
                row = new Row()
                {
                    Entry1 = range2[0].ToString(),
                    Entry2 = range2[1].ToString(),
                    Entry3 = range2[2].ToString(),
                    Entry4 = range2[3].ToString(),
                    Entry5 = range2[4].ToString(),
                    Entry6 = range2[5].ToString(),
                    Entry7 = range2[6].ToString(),
                    Entry8 = range2[7].ToString(),
                    Entry9 = range2[8].ToString(),
                    Entry10 = range2[9].ToString(),
                    Entry11 = range2[10].ToString(),
                    Entry12 = range2[11].ToString(),
                    Entry13 = range2[12].ToString(),
                    Entry14 = range2[13].ToString(),
                    Entry15 = range2[14].ToString(),
                    Entry16 = range2[15].ToString(),
                    Entry17 = range2[16].ToString(),
                    Entry18 = range2[17].ToString(),
                    Entry19 = range2[18].ToString(),
                    Entry20 = range2[19].ToString(),
                    Entry21 = range2[20].ToString(),
                    Entry22 = range2[21].ToString(),
                    Entry23 = range2[22].ToString(),
                    Entry24 = range2[23].ToString(),
                    Entry25 = range2[24].ToString(),
                    Entry26 = range2[25].ToString(),
                    Entry27 = range2[26].ToString(),
                    Entry28 = range2[27].ToString(),
                    Entry29 = range2[28].ToString(),
                    Entry30 = range2[29].ToString()
                };
                allRows2.Add(row);
            }

            //for article RU Report
            if (isRUReport) 
            {
                foreach (var item1 in allRows1)
                {
                    newSheet[$"A{rowNum}"].Value = item1.Entry1;
                    newSheet[$"B{rowNum}"].Value = item1.Entry2;
                    newSheet[$"C{rowNum}"].Value = item1.Entry3;
                    newSheet[$"D{rowNum}"].Value = item1.Entry4;
                    newSheet[$"E{rowNum}"].Value = item1.Entry5;
                    newSheet[$"F{rowNum}"].Value = item1.Entry6;
                    newSheet[$"G{rowNum}"].Value = item1.Entry7;
                    newSheet[$"H{rowNum}"].Value = item1.Entry8;
                    newSheet[$"I{rowNum}"].Value = item1.Entry9;
                    newSheet[$"J{rowNum}"].Value = item1.Entry10;
                    newSheet[$"K{rowNum}"].Value = item1.Entry11;
                    newSheet[$"L{rowNum}"].Value = item1.Entry12;
                    newSheet[$"M{rowNum}"].Value = item1.Entry13;
                    newSheet[$"N{rowNum}"].Value = item1.Entry14;
                    newSheet[$"O{rowNum}"].Value = item1.Entry15;
                    newSheet[$"P{rowNum}"].Value = item1.Entry16;
                    newSheet[$"Q{rowNum}"].Value = item1.Entry17;
                    newSheet[$"R{rowNum}"].Value = item1.Entry18;
                    newSheet[$"S{rowNum}"].Value = item1.Entry19;
                    newSheet[$"T{rowNum}"].Value = item1.Entry20;
                    newSheet[$"U{rowNum}"].Value = item1.Entry21;
                    newSheet[$"V{rowNum}"].Value = item1.Entry22;
                    newSheet[$"W{rowNum}"].Value = item1.Entry23;
                    newSheet[$"X{rowNum}"].Value = item1.Entry24;
                    newSheet[$"Y{rowNum}"].Value = item1.Entry25;
                    newSheet[$"Z{rowNum}"].Value = item1.Entry26;
                    newSheet[$"AA{rowNum}"].Value = item1.Entry27;
                    newSheet[$"AB{rowNum}"].Value = item1.Entry28;
                    newSheet[$"AC{rowNum}"].Value = item1.Entry29;
                    newSheet[$"AD{rowNum}"].Value = item1.Entry30;
                    isNew = 1;
                    foreach (var item2 in allRows2)
                    {
                        //comparing article number and range , same article no. exists in different range
                        if (item1.Entry1.Equals(item2.Entry1) && item1.Entry16.Equals(item2.Entry16))
                        {
                            isNew = 0;
                            if (!item1.Equals(item2))
                            {
                                Console.WriteLine("rows did not match for article no. " + item1.Entry1 + " range : " + item1.Entry3);

                                if (item1.Entry2 != item2.Entry2)
                                {
                                    newSheet[$"B{rowNum}"].Value += "->" + item2.Entry2;
                                    newSheet[$"B{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry3 != item2.Entry3)
                                {
                                    newSheet[$"C{rowNum}"].Value += "->" + item2.Entry3;
                                    newSheet[$"C{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry4 != item2.Entry4)
                                {
                                    newSheet[$"D{rowNum}"].Value += "->" + item2.Entry4;
                                    newSheet[$"D{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry5 != item2.Entry5)
                                {
                                    newSheet[$"E{rowNum}"].Value += "->" + item2.Entry5;
                                    newSheet[$"E{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry6 != item2.Entry6)
                                {
                                    newSheet[$"F{rowNum}"].Value += "->" + item2.Entry6;
                                    newSheet[$"F{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry7 != item2.Entry7)
                                {
                                    newSheet[$"G{rowNum}"].Value += "->" + item2.Entry7;
                                    newSheet[$"G{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry8 != item2.Entry8)
                                {
                                    newSheet[$"H{rowNum}"].Value += "->" + item2.Entry8;
                                    newSheet[$"H{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry9 != item2.Entry9)
                                {
                                    newSheet[$"I{rowNum}"].Value += "->" + item2.Entry9;
                                    newSheet[$"I{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry10 != item2.Entry10)
                                {
                                    newSheet[$"J{rowNum}"].Value += "->" + item2.Entry10;
                                    newSheet[$"J{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry11 != item2.Entry11)
                                {
                                    newSheet[$"K{rowNum}"].Value += "->" + item2.Entry11;
                                    newSheet[$"K{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry12 != item2.Entry12)
                                {
                                    newSheet[$"L{rowNum}"].Value += "->" + item2.Entry12;
                                    newSheet[$"L{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry13 != item2.Entry13)
                                {
                                    newSheet[$"M{rowNum}"].Value += "->" + item2.Entry13;
                                    newSheet[$"M{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry14 != item2.Entry14)
                                {
                                    newSheet[$"N{rowNum}"].Value += "->" + item2.Entry14;
                                    newSheet[$"N{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry15 != item2.Entry15)
                                {
                                    newSheet[$"O{rowNum}"].Value += "->" + item2.Entry15;
                                    newSheet[$"O{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry16 != item2.Entry16)
                                {
                                    newSheet[$"P{rowNum}"].Value += "->" + item2.Entry16;
                                    newSheet[$"P{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry17 != item2.Entry17)
                                {
                                    newSheet[$"Q{rowNum}"].Value += "->" + item2.Entry17;
                                    newSheet[$"Q{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry18 != item2.Entry18)
                                {
                                    newSheet[$"R{rowNum}"].Value += "->" + item2.Entry18;
                                    newSheet[$"R{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry19 != item2.Entry19)
                                {
                                    newSheet[$"S{rowNum}"].Value += "->" + item2.Entry19;
                                    newSheet[$"S{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry20 != item2.Entry20)
                                {
                                    newSheet[$"T{rowNum}"].Value += "->" + item2.Entry20;
                                    newSheet[$"T{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry21 != item2.Entry21)
                                {
                                    newSheet[$"U{rowNum}"].Value += "->" + item2.Entry21;
                                    newSheet[$"U{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry22 != item2.Entry22)
                                {
                                    newSheet[$"V{rowNum}"].Value += "->" + item2.Entry22;
                                    newSheet[$"V{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry23 != item2.Entry23)
                                {
                                    newSheet[$"W{rowNum}"].Value += "->" + item2.Entry23;
                                    newSheet[$"W{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry24 != item2.Entry24)
                                {
                                    newSheet[$"X{rowNum}"].Value += "->" + item2.Entry24;
                                    newSheet[$"X{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry25 != item2.Entry25)
                                {
                                    newSheet[$"Y{rowNum}"].Value += "->" + item2.Entry25;
                                    newSheet[$"Y{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry26 != item2.Entry26)
                                {
                                    newSheet[$"Z{rowNum}"].Value += "->" + item2.Entry26;
                                    newSheet[$"Z{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry27 != item2.Entry27)
                                {
                                    newSheet[$"AA{rowNum}"].Value += "->" + item2.Entry27;
                                    newSheet[$"AA{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry28 != item2.Entry28)
                                {
                                    newSheet[$"AB{rowNum}"].Value += "->" + item2.Entry28;
                                    newSheet[$"AB{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry29 != item2.Entry29)
                                {
                                    newSheet[$"AC{rowNum}"].Value += "->" + item2.Entry29;
                                    newSheet[$"AC{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }
                                if (item1.Entry30 != item2.Entry30)
                                {
                                    newSheet[$"AD{rowNum}"].Value += "->" + item2.Entry30;
                                    newSheet[$"AD{rowNum}"].Style.BackgroundColor = "#ff8080";
                                }

                                //rowNum++;
                            }

                            break;
                        }
                    }
                    if (isNew == 1)
                    {
                        entirelyNewRecord++;
                        newSheet[$"A{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"B{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"C{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"D{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"E{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"F{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"G{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"H{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"I{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"J{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"K{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"L{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"M{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"N{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"O{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"P{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"Q{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"R{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"S{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"T{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"U{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"V{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"W{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"X{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"Y{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"Z{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AA{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AB{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AC{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AD{rowNum}"].Style.BackgroundColor = "#99ff66";
                    }


                    rowNum++;
                }

                Console.WriteLine("\nSearching for deleted records(if any)");
                
                //writing records deleted after baseload
                foreach (var item2 in allRows2)
                {
                    flag = false;
                    foreach (var item1 in allRows1)
                    {
                        if (item2.Entry1.Equals(item1.Entry1) && item2.Entry16.Equals(item1.Entry16))
                        {
                            flag = true;
                            break;
                        }
                    }

                    if (!flag)
                    {
                        temp++;
                        //write the record in the sheet
                        newSheet[$"A{rowNum}"].Value = item2.Entry1;
                        newSheet[$"B{rowNum}"].Value = item2.Entry2;
                        newSheet[$"C{rowNum}"].Value = item2.Entry3;
                        newSheet[$"D{rowNum}"].Value = item2.Entry4;
                        newSheet[$"E{rowNum}"].Value = item2.Entry5;
                        newSheet[$"F{rowNum}"].Value = item2.Entry6;
                        newSheet[$"G{rowNum}"].Value = item2.Entry7;
                        newSheet[$"H{rowNum}"].Value = item2.Entry8;
                        newSheet[$"I{rowNum}"].Value = item2.Entry9;
                        newSheet[$"J{rowNum}"].Value = item2.Entry10;
                        newSheet[$"K{rowNum}"].Value = item2.Entry11;
                        newSheet[$"L{rowNum}"].Value = item2.Entry12;
                        newSheet[$"M{rowNum}"].Value = item2.Entry13;
                        newSheet[$"N{rowNum}"].Value = item2.Entry14;
                        newSheet[$"O{rowNum}"].Value = item2.Entry15;
                        newSheet[$"P{rowNum}"].Value = item2.Entry16;
                        newSheet[$"Q{rowNum}"].Value = item2.Entry17;
                        newSheet[$"R{rowNum}"].Value = item2.Entry18;
                        newSheet[$"S{rowNum}"].Value = item2.Entry19;
                        newSheet[$"T{rowNum}"].Value = item2.Entry20;
                        newSheet[$"U{rowNum}"].Value = item2.Entry21;
                        newSheet[$"V{rowNum}"].Value = item2.Entry22;
                        newSheet[$"W{rowNum}"].Value = item2.Entry23;
                        newSheet[$"X{rowNum}"].Value = item2.Entry24;
                        newSheet[$"Y{rowNum}"].Value = item2.Entry25;
                        newSheet[$"Z{rowNum}"].Value = item2.Entry26;
                        newSheet[$"AA{rowNum}"].Value = item2.Entry27;
                        newSheet[$"AB{rowNum}"].Value = item2.Entry28;
                        newSheet[$"AC{rowNum}"].Value = item2.Entry29;
                        newSheet[$"AD{rowNum}"].Value = item2.Entry30;

                        //color the row
                        newSheet[$"A{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"B{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"C{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"D{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"E{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"F{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"G{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"H{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"I{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"J{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"K{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"L{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"M{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"N{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"O{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"P{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"Q{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"R{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"S{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"T{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"U{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"V{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"W{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"X{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"Y{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"Z{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AA{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AB{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AC{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AD{rowNum}"].Style.BackgroundColor = "#ffff00";

                        rowNum++;
                    }
                }
            }
            //for Article Report
            else 
            {
                foreach (var item1 in allRows1)
                {
                    newSheet[$"A{rowNum}"].Value = item1.Entry1;
                    newSheet[$"B{rowNum}"].Value = item1.Entry2;
                    newSheet[$"C{rowNum}"].Value = item1.Entry3;
                    newSheet[$"D{rowNum}"].Value = item1.Entry4;
                    newSheet[$"E{rowNum}"].Value = item1.Entry5;
                    newSheet[$"F{rowNum}"].Value = item1.Entry6;
                    newSheet[$"G{rowNum}"].Value = item1.Entry7;
                    newSheet[$"H{rowNum}"].Value = item1.Entry8;
                    newSheet[$"I{rowNum}"].Value = item1.Entry9;
                    newSheet[$"J{rowNum}"].Value = item1.Entry10;
                    newSheet[$"K{rowNum}"].Value = item1.Entry11;
                    newSheet[$"L{rowNum}"].Value = item1.Entry12;
                    newSheet[$"M{rowNum}"].Value = item1.Entry13;
                    newSheet[$"N{rowNum}"].Value = item1.Entry14;
                    newSheet[$"O{rowNum}"].Value = item1.Entry15;
                    newSheet[$"P{rowNum}"].Value = item1.Entry16;
                    newSheet[$"Q{rowNum}"].Value = item1.Entry17;
                    newSheet[$"R{rowNum}"].Value = item1.Entry18;
                    newSheet[$"S{rowNum}"].Value = item1.Entry19;
                    newSheet[$"T{rowNum}"].Value = item1.Entry20;
                    newSheet[$"U{rowNum}"].Value = item1.Entry21;
                    newSheet[$"V{rowNum}"].Value = item1.Entry22;
                    newSheet[$"W{rowNum}"].Value = item1.Entry23;
                    newSheet[$"X{rowNum}"].Value = item1.Entry24;
                    newSheet[$"Y{rowNum}"].Value = item1.Entry25;
                    newSheet[$"Z{rowNum}"].Value = item1.Entry26;
                    newSheet[$"AA{rowNum}"].Value = item1.Entry27;
                    newSheet[$"AB{rowNum}"].Value = item1.Entry28;
                    newSheet[$"AC{rowNum}"].Value = item1.Entry29;
                    newSheet[$"AD{rowNum}"].Value = item1.Entry30;
                    isNew = 1;
                    foreach (var item2 in allRows2)
                    {
                        //comparing article number and range , same article no. exists in different range
                        if (item1.Entry1.Equals(item2.Entry1) && item1.Entry8.Equals(item2.Entry8))
                        {
                            isNew = 0;
                            if (!item1.Equals(item2))
                            {
                                Console.WriteLine("rows did not match for article no. " + item1.Entry1 + " range : " + item1.Entry3);

                                if (item1.Entry2 != item2.Entry2) {
                                    newSheet[$"B{rowNum}"].Value += "->" + item2.Entry2;
                                    newSheet[$"B{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry3 != item2.Entry3) {
                                    newSheet[$"C{rowNum}"].Value += "->" + item2.Entry3;
                                    newSheet[$"C{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry4 != item2.Entry4) {
                                    newSheet[$"D{rowNum}"].Value += "->" + item2.Entry4;
                                    newSheet[$"D{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry5 != item2.Entry5) {
                                    newSheet[$"E{rowNum}"].Value += "->" + item2.Entry5;
                                    newSheet[$"E{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry6 != item2.Entry6) {
                                    newSheet[$"F{rowNum}"].Value += "->" + item2.Entry6;
                                    newSheet[$"F{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry7 != item2.Entry7) {
                                    newSheet[$"G{rowNum}"].Value += "->" + item2.Entry7;
                                    newSheet[$"G{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry8 != item2.Entry8) {
                                    newSheet[$"H{rowNum}"].Value += "->" + item2.Entry8;
                                    newSheet[$"H{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry9 != item2.Entry9) {
                                    newSheet[$"I{rowNum}"].Value += "->" + item2.Entry9;
                                    newSheet[$"I{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry10 != item2.Entry10) {
                                    newSheet[$"J{rowNum}"].Value += "->" + item2.Entry10;
                                    newSheet[$"J{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry11 != item2.Entry11) {
                                    newSheet[$"K{rowNum}"].Value += "->" + item2.Entry11;
                                    newSheet[$"K{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry12 != item2.Entry12) {
                                    newSheet[$"L{rowNum}"].Value += "->" + item2.Entry12;
                                    newSheet[$"L{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry13 != item2.Entry13) {
                                    newSheet[$"M{rowNum}"].Value += "->" + item2.Entry13;
                                    newSheet[$"M{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry14 != item2.Entry14) {
                                    newSheet[$"N{rowNum}"].Value += "->" + item2.Entry14;
                                    newSheet[$"N{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry15 != item2.Entry15) {
                                    newSheet[$"O{rowNum}"].Value += "->" + item2.Entry15;
                                    newSheet[$"O{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry16 != item2.Entry16) {
                                    newSheet[$"P{rowNum}"].Value += "->" + item2.Entry16;
                                    newSheet[$"P{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry17 != item2.Entry17) {
                                    newSheet[$"Q{rowNum}"].Value += "->" + item2.Entry17;
                                    newSheet[$"Q{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry18 != item2.Entry18) {
                                    newSheet[$"R{rowNum}"].Value += "->" + item2.Entry18;
                                    newSheet[$"R{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry19 != item2.Entry19) {
                                    newSheet[$"S{rowNum}"].Value += "->" + item2.Entry19;
                                    newSheet[$"S{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry20 != item2.Entry20) {
                                    newSheet[$"T{rowNum}"].Value += "->" + item2.Entry20;
                                    newSheet[$"T{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry21 != item2.Entry21) {
                                    newSheet[$"U{rowNum}"].Value += "->" + item2.Entry21;
                                    newSheet[$"U{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry22 != item2.Entry22) {
                                    newSheet[$"V{rowNum}"].Value += "->" + item2.Entry22;
                                    newSheet[$"V{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry23 != item2.Entry23) {
                                    newSheet[$"W{rowNum}"].Value += "->" + item2.Entry23;
                                    newSheet[$"W{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry24 != item2.Entry24) {
                                    newSheet[$"X{rowNum}"].Value += "->" + item2.Entry24;
                                    newSheet[$"X{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry25 != item2.Entry25) {
                                    newSheet[$"Y{rowNum}"].Value += "->" + item2.Entry25;
                                    newSheet[$"Y{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry26 != item2.Entry26) {
                                    newSheet[$"Z{rowNum}"].Value += "->" + item2.Entry26;
                                    newSheet[$"Z{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry27 != item2.Entry27) {
                                    newSheet[$"AA{rowNum}"].Value += "->" + item2.Entry27;
                                    newSheet[$"AA{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry28 != item2.Entry28) {
                                    newSheet[$"AB{rowNum}"].Value += "->" + item2.Entry28;
                                    newSheet[$"AB{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry29 != item2.Entry29) {
                                    newSheet[$"AC{rowNum}"].Value += "->" + item2.Entry29;
                                    newSheet[$"AC{rowNum}"].Style.BackgroundColor = "#ff8080"; }
                                if (item1.Entry30 != item2.Entry30) {
                                    newSheet[$"AD{rowNum}"].Value += "->" + item2.Entry30;
                                    newSheet[$"AD{rowNum}"].Style.BackgroundColor = "#ff8080"; }

                                //rowNum++;
                            }

                            break;
                        }
                    }
                    if (isNew == 1)
                    {
                        entirelyNewRecord++;
                        newSheet[$"A{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"B{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"C{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"D{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"E{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"F{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"G{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"H{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"I{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"J{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"K{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"L{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"M{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"N{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"O{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"P{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"Q{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"R{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"S{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"T{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"U{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"V{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"W{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"X{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"Y{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"Z{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AA{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AB{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AC{rowNum}"].Style.BackgroundColor = "#99ff66";
                        newSheet[$"AD{rowNum}"].Style.BackgroundColor = "#99ff66";
                    }


                    rowNum++;
                }

                Console.WriteLine("\nSearching for deleted records(if any)");

                //writing records deleted after baseload
                foreach (var item2 in allRows2)
                {
                    flag = false;
                    foreach (var item1 in allRows1)
                    {
                        if (item1.Entry1.Equals(item2.Entry1) && item1.Entry8.Equals(item2.Entry8))
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        temp++;
                        //write the record in the sheet
                        newSheet[$"A{rowNum}"].Value = item2.Entry1;
                        newSheet[$"B{rowNum}"].Value = item2.Entry2;
                        newSheet[$"C{rowNum}"].Value = item2.Entry3;
                        newSheet[$"D{rowNum}"].Value = item2.Entry4;
                        newSheet[$"E{rowNum}"].Value = item2.Entry5;
                        newSheet[$"F{rowNum}"].Value = item2.Entry6;
                        newSheet[$"G{rowNum}"].Value = item2.Entry7;
                        newSheet[$"H{rowNum}"].Value = item2.Entry8;
                        newSheet[$"I{rowNum}"].Value = item2.Entry9;
                        newSheet[$"J{rowNum}"].Value = item2.Entry10;
                        newSheet[$"K{rowNum}"].Value = item2.Entry11;
                        newSheet[$"L{rowNum}"].Value = item2.Entry12;
                        newSheet[$"M{rowNum}"].Value = item2.Entry13;
                        newSheet[$"N{rowNum}"].Value = item2.Entry14;
                        newSheet[$"O{rowNum}"].Value = item2.Entry15;
                        newSheet[$"P{rowNum}"].Value = item2.Entry16;
                        newSheet[$"Q{rowNum}"].Value = item2.Entry17;
                        newSheet[$"R{rowNum}"].Value = item2.Entry18;
                        newSheet[$"S{rowNum}"].Value = item2.Entry19;
                        newSheet[$"T{rowNum}"].Value = item2.Entry20;
                        newSheet[$"U{rowNum}"].Value = item2.Entry21;
                        newSheet[$"V{rowNum}"].Value = item2.Entry22;
                        newSheet[$"W{rowNum}"].Value = item2.Entry23;
                        newSheet[$"X{rowNum}"].Value = item2.Entry24;
                        newSheet[$"Y{rowNum}"].Value = item2.Entry25;
                        newSheet[$"Z{rowNum}"].Value = item2.Entry26;
                        newSheet[$"AA{rowNum}"].Value = item2.Entry27;
                        newSheet[$"AB{rowNum}"].Value = item2.Entry28;
                        newSheet[$"AC{rowNum}"].Value = item2.Entry29;
                        newSheet[$"AD{rowNum}"].Value = item2.Entry30;

                        //color the row
                        newSheet[$"A{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"B{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"C{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"D{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"E{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"F{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"G{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"H{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"I{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"J{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"K{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"L{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"M{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"N{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"O{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"P{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"Q{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"R{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"S{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"T{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"U{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"V{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"W{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"X{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"Y{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"Z{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AA{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AB{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AC{rowNum}"].Style.BackgroundColor = "#ffff00";
                        newSheet[$"AD{rowNum}"].Style.BackgroundColor = "#ffff00";

                        rowNum++;
                    }
                }
            }
            

            Console.WriteLine("\nNo. of new records found after baseload : " + entirelyNewRecord);
            Console.WriteLine("\nNo. of Records deleted after baseload : " + temp);

            workBook.SaveAs("Comparison3.xlsx");
        }

        public static void CompareExcelRowwise(string path1, string path2, string rows,string column, string sheet)
        {
            Console.WriteLine("\nProcessing...\n");
            var workbook1 = WorkBook.Load(path1);
            var workbook2 = WorkBook.Load(path2);

            var worksheet1 = workbook1.GetWorkSheet(sheet); //mention name of sheet e.g. Sheet1
            var worksheet2 = workbook2.GetWorkSheet(sheet);            

            for (int i = 2; i <= Convert.ToInt32(rows); i++)
            {
                var range1 = worksheet1[$"A{i}:{column}{i}"];
                var range2 = worksheet2[$"A{i}:{column}{i}"];
                var articleNum = worksheet1[$"A{i}"].First();

                var bothRanges = range1.Zip(range2, (x, y) => new { r1 = x, r2 = y });

                foreach (var range in bothRanges)
                {
                    if (!range.r1.Value.ToString().Equals(range.r2.Value.ToString()))
                        Console.WriteLine("Rows did not match for article " +articleNum);
                    //else
                    //{
                    //    Console.WriteLine(item.r1.Value.ToString()+" : "+item.r2.Value.ToString());
                    //    Console.WriteLine("All good");
                    //}
                       
                }
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("This program compares two Excel sheets : " +
                "\n1.Compare Article Report" +
                "\n2.Compare Article RU Report"+
                "\nEnter 1 OR 2");

            int choice = Convert.ToInt32(Console.ReadLine());
            string excelPath = @"C:\Dev\ExcelComparison\ExcelComparison\bin\Debug\netcoreapp3.1";
            switch (choice)
            {
                case 1:
                    Console.WriteLine("Enter full path of Excel 1(New Excel after baseload)");
                    string path1 = Console.ReadLine();
                    Console.WriteLine("Enter full path of Excel 2");
                    string path2 = Console.ReadLine();
                    Console.WriteLine("Enter name of sheet");
                    string sheet = Console.ReadLine();
                    Console.WriteLine("Enter number of rows (including header) in sheet for excel 1");
                    string rows1 = Console.ReadLine();
                    Console.WriteLine("Enter number of rows (including header) in sheet for excel 2");
                    string rows2 = Console.ReadLine();
                    Console.WriteLine("Enter last column(for ex. D or Z or AD) for sheet 1");
                    string column1 = Console.ReadLine();
                    Console.WriteLine("Enter last column(for ex. D or Z or AD) for sheet 2");
                    string column2 = Console.ReadLine();
                    CompareExcel(path1, path2, rows1, column1, rows2, column2, sheet, false);
                    Console.WriteLine("\nSaved Excel path - " + excelPath);
                    break;
                case 2:
                    Console.WriteLine("Enter full path of Excel 1(New Excel after baseload)");
                    string path_1 = Console.ReadLine();
                    Console.WriteLine("Enter full path of Excel 2");
                    string path_2 = Console.ReadLine();
                    Console.WriteLine("Enter name of sheet");
                    string sheet_ = Console.ReadLine();
                    Console.WriteLine("Enter number of rows (including header) in sheet for excel 1");
                    string rows1_ = Console.ReadLine();
                    Console.WriteLine("Enter number of rows (including header) in sheet for excel 2");
                    string rows2_ = Console.ReadLine();
                    Console.WriteLine("Enter last column(for ex. D or Z or AD) for sheet 1");
                    string column1_ = Console.ReadLine();
                    Console.WriteLine("Enter last column(for ex. D or Z or AD) for sheet 2");
                    string column2_ = Console.ReadLine();
                    CompareExcel(path_1, path_2, rows1_, column1_, rows2_, column2_, sheet_, true);
                    Console.WriteLine("\nSaved Excel path - " + excelPath);
                    break;
                default:
                    Console.WriteLine("Enter 1 or 2");
                    break;
            }
        }
    }
}