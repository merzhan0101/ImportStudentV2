using GeneratorDiplom.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPOI.POIFS.FileSystem;
using OfficeLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GeneratorDiplom.Controllers
{
    public class GradeController : Controller
    {
        private readonly AppContext _context;
        private readonly IWebHostEnvironment _appEnvironment;


        public GradeController(AppContext context, IWebHostEnvironment appEnvironment)
        {
            _context = context;
            _appEnvironment = appEnvironment;
        }

        [HttpGet("{controller}/{studentId}/{action}")]
        public async Task<IActionResult> Info([FromRoute] int studentId)
        {
            StudentModel student = await _context.Students
                .Include(p=>p.Grades)
                .FirstAsync(p=>p.Id == studentId);

            List<SubjectModel> disciplines = await _context.Subjects
                .AsNoTracking()
                .Where(p=>p.GroupId == student.GroupId)
                .Include(p=>p.Title)
                .ToListAsync();


            if(student.Grades.Count != disciplines.Count)
            {
                foreach (SubjectModel discipline in disciplines)
                {
                    GradeModel grade = student.Grades.FirstOrDefault(p=>p.SubjectId == discipline.Id);
                    if(grade == null)
                    {
                        student.Grades.Add(new GradeModel()
                        {
                            Student = student,
                            Subject = discipline
                        });
                    }
                }
                await _context.SaveChangesAsync();
            }

            List<GradeModel> list = await _context.Grades
                .AsNoTracking()
                .Where(p => p.StudentId == studentId)
                .Include(p => p.Subject)
                .ThenInclude(p => p.Title)
                .ToListAsync();

            return View(list);
        }

        [HttpPost]
        public async Task<IActionResult> Set(int id,int val)
        {
            try
            {
                if(val < 0 || val > 5)
                {
                    return Ok("Диапозон от 0 до 5");
                }
                GradeModel grade = await _context.Grades
                    .FirstAsync(p => p.Id == id);

                grade.Score = val.ToString();

                await _context.SaveChangesAsync();

                return Ok("Успешно сохранено");
            }
            catch
            {
                return BadRequest();
            }
        }

        //[HttpGet("{controller}/{action}/{groupId}")]
        //public async Task<IActionResult> DownloadAll([FromRoute] int groupId)
        //{
        //    var students = await _context.Students
        //        .AsNoTracking()
        //            .Include(p => p.Initials)
        //            .Include(p => p.Initials_Dat)
        //            .Include(p => p.Group)
        //            .ThenInclude(p => p.Title)
        //            .Include(p => p.Group)
        //            .ThenInclude(p => p.Qualification)
        //            .Include(p => p.Grades)
        //            .ThenInclude(p => p.Subject)
        //            .ThenInclude(p => p.Title)
        //        .Where(p => p.GroupId == groupId)
        //        .ToListAsync();

        //    foreach (var student in students)
        //    {
        //        string path = Path.Combine(_appEnvironment.WebRootPath, "template.xls");
        //        string newPath = Path.Combine(_appEnvironment.WebRootPath, "files", $"{Guid.NewGuid()}.xls");


        //        byte[] bytes = await System.IO.File.ReadAllBytesAsync(path);

        //        using MemoryStream stream = new MemoryStream(bytes);

        //        ExcelOldX excel = new ExcelOldX(stream);

        //        Draw(excel, student);

        //        using FileStream fs = new FileStream(newPath, FileMode.OpenOrCreate);

        //        excel.Document.Write(fs);
        //    }
        //}

        [HttpGet("{controller}/{action}/{studentId}")]
        public async Task<IActionResult> Download([FromRoute] int studentId)
        {
            try
            {
                

                
                StudentModel student = await _context.Students
                    .AsNoTracking()
                    .Include(p=>p.Initials)
                    .Include(p=>p.Initials_Dat)
                    .Include(p => p.Group)
                    .ThenInclude(p => p.Title)
                    .Include(p => p.Group)
                    .ThenInclude(p => p.Qualification)
                    .Include(p => p.Grades)
                    .ThenInclude(p => p.Subject)
                    .ThenInclude(p => p.Title)
                    .FirstAsync(p => p.Id == studentId);

                string path = Path.Combine(_appEnvironment.WebRootPath, "template.xls");
                string newPath = Path.Combine(_appEnvironment.WebRootPath, "files", $"{Guid.NewGuid()}.xls");


                byte[] bytes = await System.IO.File.ReadAllBytesAsync(path);

                using MemoryStream stream = new MemoryStream(bytes);

                ExcelOldX excel = new ExcelOldX(stream);

                Draw(excel,student);

                using FileStream fs = new FileStream(newPath, FileMode.OpenOrCreate);

                excel.Document.Write(fs);

                return PhysicalFile(newPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{student.Initials.Get}.xls");
            }
            catch(Exception err)
            {
                return Ok(err.Message);
            }
            finally
            {
            }
        }

        public void Draw(ExcelOldX excel,StudentModel student)
        {
            var initialRu = student.Initials_Dat.Title_RU.Split(" ");
            var initialKz = student.Initials_Dat.Title_KZ.Split(" ");
            excel.Write(5, 4, student.NumApplication, 1);
            excel.Write(5, 4, student.NumApplication, 3);

            excel.Write(2, 49, initialRu[0], 6);
            excel.Write(2, 12, initialKz[0], 5);

            //excel.Write(4, 46, student.Group.StartStudies, 6);

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

            excel.Write(4, 10, student.Group.StartStudies, 5);
            excel.Write(4, 46, student.Group.StartStudies, 6);

            excel.Write(6, 9, student.Group.EndStudies, 5);
            excel.Write(8, 39, student.Group.EndStudies, 6);

            excel.Write(14, 41, student.Group.StartStudies, 6); // какая дата и где кз

            excel.Write(7, 3, student.Initials_Dat.Title_RU, 1);
            excel.Write(7, 3, student.Initials_Dat.Title_KZ, 3);

            excel.Write(9, 3, student.Group.StartStudies, 1);
            excel.Write(9, 3, student.Group.StartStudies, 3);
            excel.Write(9, 6, student.Group.EndStudies, 1);
            excel.Write(9, 6, student.Group.EndStudies, 3);

            string title_ru = "Высший колледж НАО \"Торайгыров университет\"";
            string title_kz = "\"Торайгыров университеті\"";
            string title_kz_second = "КЕАҚ жоғары колледжінің";

            excel.Write(10, 3, title_ru, 1);
            excel.Write(10, 3, title_kz, 3);
            excel.Write(6, 16, title_kz, 5);
            excel.Write(7, 6, title_kz_second, 5);
            excel.Write(9, 38, title_ru, 6);

            excel.Write(12, 2, $"{student.Group.Code} \"{student.Group.Title.Title_RU}\"", 1);
            excel.Write(12, 2, $"{student.Group.Code} \"{student.Group.Title.Title_KZ}\"", 3);

            for (int i = 0; i < 2; i++)
            {
                int row = 29;
                int cell = 2;

                int padding = 2;

                int numPage = i == 0 ? 1 : 3;

                int numSubject = 1;

                foreach (var grade in student.Grades.OrderBy(p=>p.Subject.Id))
                {
                    if (row >= 51)
                    {
                        if (cell != 9)
                        {
                            cell = 9;
                            row = 4;
                            padding = 3;
                        }
                        else
                        {
                            row = 4;
                            cell = 2;       
                            numPage++;
                            padding = 2;
                        }
                    }

                    string title = i == 0 ? grade.Subject.Title.Title_RU : grade.Subject.Title.Title_KZ;

                    excel.Write(row, cell, numSubject, numPage);
                    excel.Write(row, cell + 1, title, numPage);
                    excel.Write(row, cell + 2, grade.Subject.Hours, numPage);

                    string score = grade.Score switch
                    {
                        "5" => i == 0 ? "5 (отлично)" : "5 (үздік)",
                        "4" => i == 0 ? "4 (хорошо)" : "4 (жақсы)",
                        "3" => i == 0 ? "3 (удовл)" : "3 (қанағат)",
                        _ => grade.Score,
                    };

                    excel.Write(row, cell + 3 + padding, score, numPage);

                    if (title.Length > 37)
                    {
                        excel.Merge(row, cell + 3 + padding, row + 1, cell + 3 + padding, numPage);
                        excel.Merge(row, cell + 2, row + 1, cell + 2, numPage);
                        excel.Merge(row, cell + 1, row + 1, cell + 1, numPage);
                        excel.Merge(row, cell    , row + 1, cell    , numPage);
                        row++;
                    }

                    numSubject++;
                    row += 2;
                }
            }


            //строка,номер,название,колво часов,оценка,страница,широкая ли строка
            List<int> intIndex = new List<int>();
            int[,] gradeIndex = {
                { 29,2,3,4,7,1,0 },
                { 31,2,3,4,7,1,1 },
                { 33,2,3,4,7,1,1 },
                { 35,2,3,4,7,1,0 },
                { 37,2,3,4,7,1,1 },
                { 39,2,3,4,7,1,0 },
                { 41,2,3,4,7,1,0 },
                { 43,2,3,4,7,1,1 },
                { 45,2,3,4,7,1,0 },
                { 47,2,3,4,7,1,0 },
                { 49,2,3,4,7,1,0 },
                { 51,2,3,4,7,1,0 },
                { 4,9,10,11,15,1,0 },
                { 5,9,10,11,15,1,1 },
                { 7,9,10,11,15,1,1 },
                { 9,9,10,11,15,1,1 },
                { 11,9,10,11,15,1,0 },
                { 13,9,10,11,15,1,0 },
                { 15,9,10,11,15,1,1 },
                { 17,9,10,11,15,1,1 },
                { 19,9,10,11,15,1,1 },
                { 21,9,10,11,15,1,1 },
                { 23,9,10,11,15,1,1 },
                { 25,9,10,11,15,1,1 },
                { 27,9,10,11,15,1,1 },
                { 29,9,10,11,15,1,0 },
                { 31,9,10,11,15,1,1 },
                { 33,9,10,11,15,1,1 },
                { 36,9,10,11,15,1,0 },
                { 38,9,10,11,15,1,0 },
                { 40,9,10,11,15,1,1 },
                { 43,9,10,11,15,1,1 },
                { 45,9,10,11,15,1,1 },
                { 48,9,10,11,15,1,0 },
                { 50,9,10,11,15,1,1 },
                { 4,2,3,4,7,2,1 },
                { 6,2,3,4,7,2,1 },
                { 8,2,3,4,7,2,1 },
                { 10,2,3,4,7,2,1 },
                { 12,2,3,4,7,2,0 },
                { 14,2,3,4,7,2,1 },
                { 16,2,3,4,7,2,1 },
                { 18,2,3,4,7,2,0 },
                { 20,2,3,4,7,2,0 }
            };

            foreach (GradeModel grade in student.Grades)
            {
                if (intIndex.Count >= gradeIndex.Length && string.IsNullOrEmpty(grade.Score))
                    break;

                if (string.IsNullOrEmpty(grade.Score) && grade.Subject == null)
                    continue;
                try
                {
                    for (int i = 0; i < gradeIndex.Length - 1; i++)
                    {
                        if (!intIndex.Contains(i) &&
                            ((grade.Subject.Title.Title_RU.Length > 33 && gradeIndex[i, 6] == 1) ||
                            (grade.Subject.Title.Title_RU.Length < 33 && gradeIndex[i, 6] == 0)))
                        {
                            intIndex.Add(i);
                            excel.Write(gradeIndex[i, 0], gradeIndex[i, 2], grade.Subject.Title.Title_RU, gradeIndex[i, 5]);
                            excel.Write(gradeIndex[i, 0], gradeIndex[i, 3], grade.Subject.Hours, gradeIndex[i, 5]);
                            excel.Write(gradeIndex[i, 0], gradeIndex[i, 4], Score(grade.Score), gradeIndex[i, 5]);

                            //break;
                        }
                    }
                }
                catch { }

            }
            try
            {
                for (int stepCount = 0; stepCount < student.Grades.Count - 1; stepCount++)
                {
                    excel.Write(gradeIndex[stepCount, 0], gradeIndex[stepCount, 1], stepCount + 1, gradeIndex[stepCount, 5]);
                }
            }
            catch { }
        }

        static string Score(string num) => num switch
        {
            "5" => "5 (отлично)",
            "4" => "4 (хорошо)",
            "3" => "3 (удовл)",
            _ => num,
        };
    }
}
