using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GeneratorDiplom;
using GeneratorDiplom.Models;
using Microsoft.EntityFrameworkCore;
using OfficeLib;

namespace ImportStudentV2
{
    internal class Program
    {
        static int GroupId { get; set; }

        static GeneratorDiplom.AppContext Repository;

        static List<StudentModel> Students => Repository.Students
            .Where(p=> GroupId != 0 && p.GroupId == GroupId)
            .Include(p => p.Initials)
            .Include(p => p.Initials_Dat)
            .Include(p => p.Group)
            .ThenInclude(p => p.Title)
            .Include(p => p.Group)
            .ThenInclude(p => p.Qualification)
            .Include(p => p.Grades)
            .ThenInclude(p => p.Subject)
            .ThenInclude(p => p.Title)
            .Include(p=>p.Topic)
            .ToList();


        static List<GroupModel> Groups => Repository.Groups
            .AsNoTracking()
            .Include(p=>p.Title)
            .ToList();


        static void Main(string[] args)
        {
            Repository = new GeneratorDiplom.AppContext();
            if (GroupId == 0)
                SelectGroup();

            Log("Выберите действие:");


            int num = ViewSelect("Темы","Ведомость","Дипломы","Рег номера","Удалить ведомость", "Удалить  все");

            switch(num+1)
            {
                case 1:
                    ImportTopic();
                    break;
                case 2:
                    ImportVedomost();
                    break;
                case 3:
                    ImportDiplom();
                    break;
                case 4:
                    ImportRegNum();
                    break;
                case 5:
                    var grades = Repository.Grades
                        .Include(p => p.Student)
                        .Where(p => p.Student.GroupId == GroupId)
                        .ToList();

                    Repository.Grades.RemoveRange(grades);

                    Repository.SaveChanges();

                    break;
                case 6:
                    var grades1 = Repository.Grades
                        .Include(p => p.Student)
                        .Where(p => p.Student.GroupId == GroupId)
                        .ToList();


                    var subjects1 = Repository.Subjects
                        .Where(p => p.GroupId == GroupId)
                        .ToList();

                    var subjects = Repository.Subjects
                        .Where(p => p.GroupId == GroupId)
                        .Select(s => s.TitleId)
                        .ToList();

                    var localizers = Repository.Localizers
                    .Where(l => subjects.Contains(l.Id))
                    .ToList();

                    Repository.Grades.RemoveRange(grades1);
                    Repository.Subjects.RemoveRange(subjects1);
                    Repository.Localizers.RemoveRange(localizers);

                    Repository.SaveChanges();

                    break;
            }

            Repository.Dispose();
        }

        static void ImportRegNum()
        {
            string path = ViewAnswer("Введите путь до файла с рег номерами");
            int rowSurname = int.Parse(ViewAnswer("Колонка с фамилиями"));
            int rowNum = int.Parse(ViewAnswer("Колонка с номерами"));
            var document = GetExcel(path);

            for (int i = 1; i <= Students.Where(p=>p.GroupId == GroupId).Count(); i++)
            {
                try
                {
                    string surname = SuperSplit(document.Read(i, rowSurname))[0];
                    var student = Students
                        .First(p => p.Initials.Title_RU.StartsWith(surname));
                    student.NumApplication = document.Read(i, rowNum);

                    Log($"{student.Initials.Get} {student.NumApplication}");
                }
                catch(Exception err)
                {
                    Log(err.Message);
                    break;
                }
            }
            Repository.SaveChanges();

        }

        static void ImportDiplom()
        {
            List<int> use_ids = new List<int>();
            string path = ViewAnswer("Введите путь до файла с дипломом:");
            //string path = @"D:\Diploma\ImportStudentV2\bin\Debug\netcoreapp3.1\123\2023\college\ЭСЖ419\template.xls";
            //var document = GetExcelOld(path);
            //int rowDiplomRU = int.Parse(ViewAnswer("Номер строки с темой диплома (русский)"));
            //int rowDiplomRU = 26;
            //int rowDiplomKZ = int.Parse(ViewAnswer("Номер строки с темой диплома (казахский)"));
            //int rowDiplomKZ = 28;
            foreach (var student in Students)
            {
                var document = GetExcelOld(path);
                //if (student.TopicId == null)
                //{
                //    Log($"{student.Initials.Title_RU} нету темы дипломной работы");
                //    continue;
                //}
                //if (string.IsNullOrEmpty(student.NumApplication))
                //{
                //    Log($"{student.Initials.Title_RU} нету номера диплома");
                //    continue;
                //}

                //byte[] bytes = System.IO.File.ReadAllBytes(path);
                //using MemoryStream stream = new MemoryStream(bytes);
                //ExcelOldX excel = new ExcelOldX(stream);



                //для диплома нужно
                //Draw(document, student);



                document.Write(5, 3, student.NumApplication, 1);
                document.Write(5, 3, student.NumApplication, 3);

                document.Write(7, 2, student.Initials.Title_RU, 1);
                document.Write(7, 2, student.Initials.Title_KZ, 3);

                document.Write(9, 3, (int)student.DateApplication, 1);
                document.Write(9, 3, (int)student.DateApplication, 3);

                document.Write(9, 6, student.Group.EndStudies, 1);
                document.Write(9, 6, student.Group.EndStudies, 3);

                string nameColegeRus = "Высший колледж НАО \"Торайгыров университет\"";
                string nameColegeKaz = "\"Торайғыров университеті\" КЕАҚ жоғары колледжі";
                document.Write(10, 2, nameColegeRus, 1);
                document.Write(10, 2, nameColegeKaz, 3);

                document.Write(12, 2, student.Group.Title.Title_RU, 1);
                document.Write(12, 2, student.Group.Title.Title_KZ, 3);

                document.Write(18, 3, student.Group.Qualification.Title_RU, 1);
                document.Write(18, 3, student.Group.Qualification.Title_KZ, 3);



                //document.Write(rowDiplomRU, 3, $"Защита дипломного проекта: {student.Topic.Title_RU}", 1);
                //document.Write(rowDiplomKZ, 3, $"Диплом жобасын қорғау: {student.Topic.Title_KZ}", 3);

                //var defend = GetSubject("Дипломный проект");

                //if (defend != null)
                //{
                //    var grad = student.Grades.FirstOrDefault(p => p.Subject == defend);
                //    if (grad != null)
                //    {
                //        var diplom_score_ru = GetScore(grad.Score, true);
                //        var diplom_score_kz = GetScore(grad.Score, false);
                //        document.Write(rowDiplomRU, 7, diplom_score_ru.Lang, 1);
                //        //document.Write(rowDiplomRU, 14, diplom_score_ru.Ball, 2);
                //        //document.Write(rowDiplomRU, 13, diplom_score_ru.Letter, 2);
                //        //document.Write(rowDiplomRU, 12, diplom_score_ru.Kda, 2);
                //        //document.Write(rowDiplomKZ, 12, diplom_score_ru.Kda, 4);
                //        //document.Write(rowDiplomKZ, 13, diplom_score_ru.Letter, 4);
                //        //document.Write(rowDiplomKZ, 14, diplom_score_ru.Ball, 4);
                //        document.Write(rowDiplomKZ, 7, diplom_score_kz.Lang, 3);
                //    }
                //}
                //else
                //{
                //    Log("Нету Оценок за защиту");
                //}



                for (int page = 1; page <= 4; page++)
                {
                    int offsetTitle = 3;
                    int offsetScore = 8;
                    for (int count = 1; count <= 2; count++)
                    {
                        if ((page == 1 || page == 3) && count == 1)
                        {
                            offsetTitle = 2;
                            offsetScore = 7;
                        }
                        else if ((page == 1 || page == 3) && count == 2)
                        {
                            offsetTitle = 10;
                            offsetScore = 15;
                        }
                        else if ((page == 2 || page == 4) && count == 1)
                        {
                            offsetTitle = 2;
                            offsetScore = 7;
                        }
                        else if ((page == 2 || page == 4) && count == 2)
                        {
                            offsetTitle = 10;
                            offsetScore = 15;
                        }
                        for (int row = 1; row <= 53; row++)
                        {
                            string title_ru = document.Read(row, offsetTitle, page);
                            //.Replace(" (факультатив)", "")
                            //.Replace(" (курстық жоба)", "")
                            //.Replace(" (курсовой проект)", "");
                            //.Replace(" (курсовая работа)", "")
                            //.Replace(" (курстық жұмыс)", "");

                            if (string.IsNullOrEmpty(title_ru))
                                continue;

                            int.TryParse(document.Read(row, offsetTitle + 1, page), out int hours);

                            //if (hours == 0)
                            //    continue;

                            var subject = GetSubject(title_ru, hours);
                            if (subject == null)
                            {
                                Log($"Предмет '{title_ru}' не найден");
                                continue;
                            }

                            var grade = student.Grades.FirstOrDefault(p => p.SubjectId == subject.Id);

                            if (grade == null || string.IsNullOrEmpty(grade.Score) || grade.Score == null)
                            {
                                Log($"[{student.Initials.Get}] Нету оценки: '{title_ru}'");
                                continue;
                            }


                            grade.Score = grade.Score.Split(',')[0];



                            bool lang = page == 1 || page == 2;

                            var score = GetScore(grade.Score, lang);

                            if (!string.IsNullOrEmpty(score.Kda))
                            {
                                document.Write(row, offsetTitle + 2, score.Kda, page);
                                document.Write(row, offsetTitle + 3, score.Letter, page);
                                document.Write(row, offsetTitle + 4, score.Ball, page);
                            }

                            document.Write(row, offsetScore, score.Lang, page);
                        }
                    }
                }

                document.Save(Path.Combine("Группы", $"{student.Initials.Title_RU}.xls"));
                Log($"Документ '{student.Initials.Title_RU}'.xls сгенерирован");
                //break;
            }
        }

        static public void Draw(ExcelOldX excel, StudentModel student)
        {
            var initialRu = student.Initials_Dat.Title_RU.Split(" ");
            var initialKz = student.Initials_Dat.Title_KZ.Split(" ");

            excel.Write(2, 49, initialRu[0], 6);
            excel.Write(2, 12, initialKz[0], 5);

            if (initialRu.Length == 3)
            {
                excel.Write(3, 38, $"{initialRu[1]} {initialRu[2]}", 6);
                excel.Write(3, 7, $"{initialKz[1]} {initialKz[2]}", 5);
            }
            else
            {
                excel.Write(3, 38, initialRu[1], 6);
                excel.Write(3, 7, initialKz[1], 5);
            }

            excel.Write(15, 7, student.Group.Title.Title_KZ, 5);
            excel.Write(15, 7, student.Group.Title.Title_RU, 6);

            excel.Write(4, 10, (int)student.DateApplication, 5);
            excel.Write(4, 46, (int)student.DateApplication, 6);

            excel.Write(6, 9, student.Group.EndStudies, 5);
            excel.Write(8, 39, student.Group.EndStudies, 6);

            //excel.Write(14, 41, student.Group.StartStudies, 6); // какая дата и где кз

            string title_ru = "Высший колледж НАО \"Торайгыров университет\"";
            string title_kz = "\"Торайгыров университеті\"";
            string title_kz_second = "КЕАҚ жоғары колледжінің";

            excel.Write(6, 16, title_kz, 5);
            excel.Write(7, 6, title_kz_second, 5);
            excel.Write(9, 38, title_ru, 6);

            string groupTitleRu = student.Group.Title.Title_RU;
            string groupTitleKz = student.Group.Title.Title_KZ;

            string[] groupTitleRuAr = groupTitleRu.Split(" ");
            string[] groupTitleKzAr = groupTitleKz.Split(" ");

            string titleRu1 = groupTitleRuAr[0] + ' ' + groupTitleRuAr[1];
            string titleRu2 = null;
            for (int i = 2; i < groupTitleRuAr.Length; i++)
            {
                titleRu2 += groupTitleRuAr[i] + ' ';
            }

            string titleKz1 = groupTitleKzAr[0] + ' ' + groupTitleKzAr[1];
            string titleKz2 = null;
            for (int i = 2; i < groupTitleKzAr.Length; i++)
            {
                titleKz2 += groupTitleKzAr[i] + ' ';
            }

            excel.Write(10, 50, titleRu1, 6);
            excel.Write(11, 38, titleRu2, 6);

            excel.Write(9, 11, titleKz1, 5);
            excel.Write(10, 6, titleKz2, 5);


            excel.Write(15, 7, student.Group.Qualification.Title_RU, 5);
            excel.Write(16, 38, student.Group.Qualification.Title_KZ, 6);

            excel.Write(24, 13, student.NumApplication, 5);
            excel.Write(24, 51, student.NumApplication, 6);


        }


        static Score GetScore(string score_text,bool isRus)
        {
            int.TryParse(score_text,out int score);

            Score s = new Score();

            

            if (score > 0)
            {
                if (score <= 100 && score >= 95)
                {
                    s.Letter = "A";
                    s.Kda = "4,00";
                }
                else if (score <= 94 && score >= 90)
                {
                    s.Letter = "A-";
                    s.Kda = "3,67";
                }
                else if (score <= 89 && score >= 85)
                {
                    s.Letter = "B+";
                    s.Kda = "3,33";
                }
                else if (score <= 84 && score >= 80)
                {
                    s.Letter = "B";
                    s.Kda = "3,00";
                }
                else if (score <= 79 && score >= 75)
                {
                    s.Letter = "B-";
                    s.Kda = "2,67";
                }
                else if (score <= 74 && score >= 70)
                {
                    s.Letter = "C+";
                    s.Kda = "2,33";
                }
                else if (score <= 69 && score >= 65)
                {
                    s.Letter = "C";
                    s.Kda = "2,00";
                }
                else if (score <= 64 && score >= 60)
                {
                    s.Letter = "C-";
                    s.Kda = "1,67";
                }
                else if(score <= 59 && score >= 55)
                {
                    s.Letter = "D+";
                    s.Kda = "1,33";
                }
                else if(score <= 54 && score >= 50)
                {
                    s.Letter = "D";
                    s.Kda = "1,00";
                }
                else
                {
                    s.Lang = score switch
                    {
                        5 => isRus ? "5 (отлично)" : "5 (үздік)",
                        4 => isRus ? "4 (хорошо)" : "4 (жақсы)",
                        3 => isRus ? "3 (удовл)" : "3 (қанағат)"
                    };
                }
            }
            else
            {
                s.Lang = score_text switch
                {
                    "зач." => isRus ? "зачет" : "сынақ",
                    "зач" => isRus ? "зачет" : "сынақ",
                    "зачет" => isRus ? "зачет" : "сынақ",
                    "сынақ" => isRus ? "зачет" : "сынақ",
                    "зачёт\r\n" => isRus ? "зачет" : "сынақ",
                    "зачёт" => isRus ? "зачет" : "сынақ",
                    _ => score_text
                };
            }


            s.Ball = score;

            if (string.IsNullOrEmpty(s.Lang))
            {
                s.Lang = s.Letter switch
                {
                    
                    "A" => isRus ? "5 (отлично)" : "5 (үздік)",
                    "A-" => isRus ? "5 (отлично)" : "5 (үздік)",
                    "B+" => isRus ? "4 (хорошо)" : "4 (жақсы)",
                    "B" => isRus ? "4 (хорошо)" : "4 (жақсы)",
                    "B-" => isRus ? "4 (хорошо)" : "4 (жақсы)",
                    "C+" => isRus ? "4 (хорошо)" : "4 (жақсы)",
                    "C" => isRus ? "3 (удовл)" : "3 (қанағат)",
                    "C-" => isRus ? "3 (удовл)" : "3 (қанағат)",
                    "D+" => isRus ? "3 (удовл)" : "3 (қанағат)",
                    "D" => isRus ? "3 (удовл)" : "3 (қанағат)",
                    _ => score_text,
                };
            }
            return s;
        }


        static void ImportTopic()
        {
            string path = ViewAnswer("Укажите путь до файла с дипломными темами:");
            var document = GetExcel(path);
            int cellTitle = int.Parse(ViewAnswer("Укажите номер колонки с фамилиями"));
            int cellNameRU = int.Parse(ViewAnswer("Укажите номер колонки с названиями дипломных проектов (на русском)"));
            int cellNameKZ = int.Parse(ViewAnswer("Укажите номер колонки с названиями дипломных проектов (на казахском)"));

            for (int row = 1; row <= Students.Count; row++)
            {
                string data = document.Read(row, cellTitle);
                if (string.IsNullOrEmpty(data))
                    continue;
                string surname = SuperSplit(data)[0];
                var student = Students.FirstOrDefault(p => p.Initials.Title_RU.StartsWith(surname));
                if(student == null)
                {
                    Log($"{surname} не найден в БД");
                    continue;
                }
                if(student.TopicId == null)
                {
                    student.Topic = new LocalizerModel()
                    {
                        Title_RU = document.Read(row, cellNameRU),
                        Title_KZ = document.Read(row, cellNameKZ),
                    };
                }
            }
            Repository.SaveChanges();
        }

        static void ImportVedomost()
        {
            string path = ViewAnswer("Укажите путь до файла ведомости");
            var document = GetExcel(path);
            bool bug = false;
            int rowTitleRU = int.Parse(ViewAnswer("Укажите номер строки где расположены названия предметов(на русском)"));
            int rowTitleKZ = int.Parse(ViewAnswer("Укажите номер строки где расположены названия предметов(на казахском)"));

            int rowHours = int.Parse(ViewAnswer("Укажите номер строки с часами"));

            int startRow = int.Parse(ViewAnswer("Введите номер строки студентов:"));
            int nameColumn = int.Parse(ViewAnswer("Введите номер колонки фамилии студентов"));

            for (int col = 1; col <= 115; col++)
            {
                string title = document.Read(rowTitleRU, col);
                if (string.IsNullOrEmpty(title))
                    continue;

                int.TryParse(document.Read(rowHours, col), out int hours);

                SubjectModel subject = GetSubject(title, hours);

                if (subject == null)
                {
                    //Log($"Создать предмет {title}?");
                    //var key = Console.ReadKey();
                    //Log("Y - чтобы создать; Escape - чтобы закончить цикл");
                    //if (key.Key == ConsoleKey.Y)
                    if(true)
                    {
                        subject = new SubjectModel()
                        {
                            GroupId = GroupId,
                            Title = new LocalizerModel()
                            {
                                Title_RU = title,
                                Title_KZ = document.Read(rowTitleKZ, col)
                            },
                            Hours = hours
                        };
                        Repository.SaveChanges();
                        Log($"Создан предмет {title}");
                    }
                    //else if (key.Key == ConsoleKey.Escape)
                    //{
                    //    break;
                    //}
                    //else
                    //{
                    //    continue;
                    //}
                }

                //if (subject.Title.Title_RU == "Алгоритмизация и программирование")
                //{
                //    if (!bug)
                //        bug = !bug;
                //    else
                //        subject = GetSubject("Алгоритмизация и программирование (курсовой проект)");
                //}
                for (int row = startRow; row < startRow + Students.Count; row++)
                {
                    string surname = document.Read(row, nameColumn);
                    StudentModel student = Students.FirstOrDefault(p=>p.Initials.Title_RU.StartsWith(surname));
                    if (student == null)
                    {
                        Log($"Студент {surname} не найден в базе данных");
                        continue;
                    }
                    GradeModel grade = student.Grades.FirstOrDefault(p => p.SubjectId == subject.Id);
                    if (grade == null)
                    {
                        grade = new GradeModel()
                        {
                            Student = student,
                            Subject = subject
                        };
                        student.Grades.Add(grade);
                    }
                    string score = document.Read(row, col);
                    if (!string.IsNullOrEmpty(score) && grade.Score != score)
                    {
                        grade.Score = score;
                    }
                }
                Log($"{subject.Title.Title_RU} оценки проставлены");
            }
            Repository.SaveChanges();
        }

        static List<SubjectModel> GetSubjects(string title, int hour = 0) => Repository.Subjects
            .Include(p => p.Title)
            .Where(p => (p.Title.Title_RU == title || p.Title.Title_KZ == title) && p.GroupId == GroupId && p.Hours == hour)
            .ToList();

        static SubjectModel GetSubject(string title,int hour = 0) =>
            Repository.Subjects
            .Include(p => p.Title)
            .FirstOrDefault(p => (p.Title.Title_RU == title || p.Title.Title_KZ == title) && p.GroupId == GroupId && p.Hours == hour);

        static ExcelOldX GetExcelOld(string path)
        {
            var excel = new ExcelOldX();
            excel.Open(path).Wait();
            return excel;
        }

        static ExcelX GetExcel(string path)
        {
            var excel = new ExcelX();
            excel.Open(path).Wait();
            return excel;
        }

        static ExcelOldX SelectExcelOld()
        {
            string path = ViewAnswer("Выберите путь до файла");

            return GetExcelOld(path);
        }

        static ExcelX SelectExcel()
        {
            string path = ViewAnswer("Выберите путь до файла");

            return GetExcel(path);
        }

        static void SelectGroup()
        {
            Log("Выберите группу:");
            GroupId = Groups[ViewSelect(Groups.Select(p => p.Title.Title_RU).ToArray())].Id;

            var group = Repository.Groups
                .AsNoTracking()
                .Include(p=>p.Title)
                .First(p=>p.Id == GroupId);

            Log($"Вы выбрали: {group.Title.Title_RU} {group.StartStudies} - {group.EndStudies}");
        }

        static string[] SuperSplit(string str)
        {
            List<string> list = new List<string>();

            var data = str.Split(' ');

            foreach (var item in data)
            {
                if (!string.IsNullOrEmpty(item))
                    list.Add(item);
            }

            return list.ToArray();
        }


        static string ViewAnswer(string question)
        {
            Log(question);
            return Console.ReadLine();
        }

        static int ViewSelect(params object[] param)
        {
            for (int i = 0; i < param.Length; i++)
            {
                Log($"{i+1}) {param[i]}");
            }
            string parseKey = Console.ReadLine();
            return int.Parse(parseKey)-1;
        }

        static void Log(string msg) => Console.WriteLine($"[{DateTime.Now.ToShortTimeString()}] {msg}");
    }
}
