using GeneratorDiplom.Models;
using GeneratorDiplom.ViewModels.Group;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GeneratorDiplom.Controllers
{
    public class GroupController : Controller
    {
        private readonly AppContext _context;

        public GroupController(AppContext context) 
            => _context = context;

        [HttpGet]
        public async Task<IActionResult> Index()
        {
            List<GroupModel> list = await _context.Groups
                .AsNoTracking()
                .Include(p=>p.Title)
                .Include(p=>p.Subjects)
                .ToListAsync();

            return View(list);
        }

        [HttpPost("{controller}/{action}/{groupId}")]
        public async Task<IActionResult> Edit([FromRoute]int groupId,GroupViewModel model)
        {
            GroupModel group = await _context.Groups
                .Include(p => p.Title)
                .Include(p => p.Qualification)
                .FirstAsync(p => p.Id == groupId);

            group.Qualification.Title_RU = model.Qualification_RU;
            group.Qualification.Title_KZ = model.Qualification_KZ;
            group.Title.Title_RU = model.Title_RU;
            group.Title.Title_KZ = model.Title_KZ;
            group.StartStudies = model.StartStudies;
            group.EndStudies = model.EndStudies;
            group.Code = model.Code;
            group.Language = model.Lang;

            await _context.SaveChangesAsync();

            return RedirectToAction("Index");
        }

        [HttpGet("{controller}/{action}/{groupId}")]
        public async Task<IActionResult> Edit([FromRoute]int groupId)
        {
            GroupModel group = await _context.Groups
                .AsNoTracking()
                .Include(p => p.Title)
                .Include(p => p.Qualification)
                .FirstAsync(p=>p.Id == groupId);

            GroupViewModel view = new GroupViewModel()
            {
                Title_RU = group.Title.Title_RU,
                Title_KZ = group.Title.Title_KZ,
                Code = group.Code,
                EndStudies = group.EndStudies,
                Lang = group.Language,
                StartStudies = group.StartStudies,
                Qualification_RU = group.Qualification.Title_RU,
                Qualification_KZ = group.Qualification.Title_KZ
            };

            return View("Create",view);
        }

        [HttpGet]
        public IActionResult Create() => View();

        [HttpPost]
        public async Task<IActionResult> Create(GroupViewModel model)
        {
            GroupModel group = new GroupModel()
            {
                Title = new LocalizerModel()
                {
                    Title_RU = model.Title_RU,
                    Title_KZ = model.Title_KZ
                },
                Qualification = new LocalizerModel()
                {
                    Title_RU = model.Qualification_RU,
                    Title_KZ = model.Qualification_KZ
                },
                Code = model.Code,
                StartStudies = model.StartStudies,
                EndStudies = model.EndStudies,
                Language = model.Lang
            };
            _context.Groups.Add(group);
            await _context.SaveChangesAsync();
            return RedirectToAction("Index");
        }

    }
}
