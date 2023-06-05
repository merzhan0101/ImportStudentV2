using GeneratorDiplom.Models;
using GeneratorDiplom.ViewModels.Student;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GeneratorDiplom.Controllers
{
    public class StudentController : Controller
    {
        private readonly AppContext _context;

        public StudentController(AppContext context) 
            => _context = context;

        [HttpGet]
        public async Task<IActionResult> Index()
        {
            List<StudentModel> listStudents = await _context.Students
                .AsNoTracking()
                .Include(p=>p.Initials)
                .Include(p=>p.Initials_Dat)
                .Include(p=>p.Group)
                .ThenInclude(p=>p.Title)
                .OrderBy(p => p.Group.Title.Title_RU)
                .Where(p=>string.IsNullOrEmpty(p.NumApplication))
                .ToListAsync();

            return View(listStudents);
        }

        [HttpGet]
        public async Task<IActionResult> Finish()
        {
            List<StudentModel> listStudents = await _context.Students
                .AsNoTracking()
                .Include(p => p.Initials)
                .Include(p=>p.Initials_Dat)
                .Include(p=>p.Grades)
                .Include(p => p.Group)
                .ThenInclude(p=>p.Subjects)
                .Include(p=>p.Group)
                .ThenInclude(p => p.Title)
                .OrderBy(p => p.Group.Title.Title_RU)
                .Where(p=>p.Grades.Count == p.Group.Subjects.Count)
                .ToListAsync();

            return View("Index", listStudents);
        }

        [HttpPost("{controller}/{action}/{studentId}")]
        public async Task<IActionResult> Edit([FromRoute]int studentId,StudentViewModel model)
        {
            StudentModel student = await _context.Students
                .Include(p=>p.Initials)
                .FirstAsync(p=>p.Id == studentId);

            student.GroupId = model.GroupId;
            student.Initials.Title_RU = $"{model.Surname_RU} {model.Name_RU} {model.Middlename_RU}";
            student.NumApplication = model.NumApplication;

            await _context.SaveChangesAsync();

            return RedirectToAction("Index");
        }

        [HttpGet("{controller}/{action}/{studentId}")]
        public async Task<IActionResult> Edit([FromRoute]int studentId)
        {
            StudentModel student = await _context.Students
                .AsNoTracking()
                .Include(p=>p.Initials)
                .Include(p=>p.Initials_Dat)
                .FirstAsync(p=>p.Id == studentId);

            List<GroupModel> groupList = await _context.Groups
                .AsNoTracking()
                .Include(p=>p.Title)
                .ToListAsync();

            var initial = student.Initials.Title_RU.Split(" ");
            var initialDatRu = student.Initials_Dat.Title_RU.Split(" ");   //null ex
            var initialDatKz = student.Initials_Dat.Title_KZ.Split(" ");

            ViewBag.Groups = new SelectList(groupList, "Id", "Title.Title_RU");

            StudentViewModel view = new StudentViewModel()
            {
                Surname_RU = initial[0],
                Middlename_RU = initial[2],
                GroupId = student.GroupId,
                NumApplication = student.NumApplication,
                Name_RU = initial[1],
                Surname_Dat_RU = initialDatRu[0],
                Name_Dat_RU = initialDatRu[1],
                Surname_Dat_KZ = initialDatKz[0],
                Name_Dat_KZ = initialDatKz[1]
            };
            if(initialDatRu.Length == 3)
            {
                view.Middlename_Dat_RU = initialDatRu[2];
                view.Middlename_Dat_KZ = initialDatKz[2];
            }

            return View("Create", view);
        }

        [HttpPost]
        public async Task<IActionResult> Create(StudentViewModel model)
        {
            StudentModel student = new StudentModel()
            {
                Initials = new LocalizerModel()
                {
                    Title_RU = $"{model.Surname_RU} {model.Name_RU} {model.Middlename_RU}"
                },
                Initials_Dat = new LocalizerModel(),
                GroupId = model.GroupId,
                NumApplication = model.NumApplication
            };

            _context.Students.Add(student);

            await _context.SaveChangesAsync();

            var subjectIds = await _context.Subjects
                .AsNoTracking()
                .Where(p => p.GroupId == student.GroupId)
                .Select(p => p.Id)
                .ToListAsync();

            foreach (int subjectId in subjectIds)
            {
                student.Grades.Add(new GradeModel()
                {
                    StudentId = student.Id,
                    Score = null,
                    SubjectId = subjectId
                });
            }

            await _context.SaveChangesAsync();

            return RedirectToAction("Info","Grade",new { studentId = student.Id });
        }

        [HttpGet]
        public async Task<IActionResult> Create()
        {
            List<GroupModel> groupList = await _context.Groups
                .AsNoTracking()
                .Include(p=>p.Title)
                .ToListAsync();

            ViewBag.Groups = new SelectList(groupList, "Id", "Title.Title_RU");

            return View(new StudentViewModel());
        }
    }
}
