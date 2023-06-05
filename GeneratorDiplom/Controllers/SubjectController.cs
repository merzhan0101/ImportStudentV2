using GeneratorDiplom.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GeneratorDiplom.Controllers
{
    public class SubjectController : Controller
    {
        private readonly AppContext _context;

        public SubjectController(AppContext context)
            => _context = context;


        [HttpGet("{controller}/{groupId}/{action}")]
        public async Task<IActionResult> List(int groupId)
        {
            List<SubjectModel> list = await _context.Subjects
                .AsNoTracking()
                .Where(p=>p.GroupId == groupId)
                .Include(p => p.Title)
                .OrderBy(p=>p.Id)
                .ToListAsync();

            return View(list);
        }

        [HttpGet("{controller}/{groupId}/{action}")]
        public IActionResult Create() => View();

        [HttpPost("{controller}/{groupId}/{action}")]
        public async Task<IActionResult> Create([FromRoute] int groupId,string title_ru,string title_kz,int hours)
        {
            SubjectModel model = new SubjectModel()
            {
                GroupId = groupId,
                Hours = hours
            };
            model.Title.Title_RU = title_ru;
            model.Title.Title_KZ = title_kz;
            _context.Subjects.Update(model);
            await _context.SaveChangesAsync();
            return RedirectToAction("List");
        }
    }
}
