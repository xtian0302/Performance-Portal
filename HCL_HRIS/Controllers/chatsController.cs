using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description; 
using HCL_HRIS.Models;

namespace HCL_HRIS.Controllers
{
    public class chatObject : chat
    {
        public string Name { get; set; }
    }
    public class chatsController : ApiController
    {
        private HCL_HRISEntities db = new HCL_HRISEntities();

        // GET: api/chats
        [HttpGet]
        public IQueryable<chat> Getchats()
        {
            int? sap_id = Int32.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            int? sup_id = leader.user_id, sup_sap_id = leader.sap_id;
            return db.chats.Where(x => (x.chat_from == usr.user_id && x.chat_to == sup_sap_id) || ((x.chat_from == sup_id && x.chat_to == usr.sap_id)));
        }
        // GET: api/chats/GetGroup
        [Route("api/chats/GetGroup")]
        [HttpGet]
        public IQueryable<chatObject> GetGroup()
        {
            int? sap_id = Int32.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            int? group_id = usr.group_id;
            List<chatObject> list = new List<chatObject>();
            foreach (var chat in db.chats.Where(x => x.group_id == group_id))
            {
                chatObject chat1 = new chatObject();
                chat1.chat_id = chat.chat_id;
                chat1.chat_from = chat.chat_from;
                chat1.chat_to = chat.chat_to;
                chat1.datetime_read = chat.datetime_read;
                chat1.chat_message = chat.chat_message;
                chat1.datetime_sent = chat.datetime_sent;
                chat1.group_id = chat.group_id;
                chat1.Name = db.users.Where(x => x.user_id == chat.chat_from).First().name;
                list.Add(chat1);
            }
            return list.AsQueryable();
        }
        [Route("api/chats/GetTL/{fromID}")]
        [HttpGet]
        public IQueryable<chatObject> GetTL(int fromID)
        {
            int toID = Int32.Parse(User.Identity.Name); 
            user tl = db.users.Where(x => x.sap_id == toID).First();
            user agent = db.users.Where(x => x.user_id == fromID).First();
            List<chatObject> list = new List<chatObject>();
            foreach (var chat in db.chats.Where(x => (x.chat_from == fromID && x.chat_to == toID) ||(x.chat_from == tl.user_id && x.chat_to == agent.sap_id)))
            {
                chatObject chat1 = new chatObject();
                chat1.chat_id = chat.chat_id;
                chat1.chat_from = chat.chat_from;
                chat1.chat_to = chat.chat_to;
                chat1.datetime_read = chat.datetime_read;
                chat1.chat_message = chat.chat_message;
                chat1.datetime_sent = chat.datetime_sent;
                chat1.group_id = chat.group_id;
                chat1.Name = db.users.Where(x => x.user_id == chat.chat_from).First().name;
                list.Add(chat1);
            }
            return list.AsQueryable();
        }
        // GET: api/chats/5
        [ResponseType(typeof(chat))]
        public async Task<IHttpActionResult> Getchat(int id)
        {
            chat chat = await db.chats.FindAsync(id);
            if (chat == null)
            {
                return NotFound();
            }

            return Ok(chat);
        }

        // PUT: api/chats/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> Putchat(int id, chat chat)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != chat.chat_id)
            {
                return BadRequest();
            }

            db.Entry(chat).State = EntityState.Modified;

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!chatExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/chats
        [ResponseType(typeof(chat))]
        public async Task<IHttpActionResult> Postchat(chat chat)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            chat.datetime_sent = DateTime.Now;
            db.chats.Add(chat);
            await db.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = chat.chat_id }, chat);
        }

        // DELETE: api/chats/5
        [ResponseType(typeof(chat))]
        public async Task<IHttpActionResult> Deletechat(int id)
        {
            chat chat = await db.chats.FindAsync(id);
            if (chat == null)
            {
                return NotFound();
            }

            db.chats.Remove(chat);
            await db.SaveChangesAsync();

            return Ok(chat);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool chatExists(int id)
        {
            return db.chats.Count(e => e.chat_id == id) > 0;
        }
    }
}