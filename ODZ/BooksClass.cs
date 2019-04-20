using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ODZ
{
    public class BooksClass
    {
        public string IDBook { get; set; }
        public string NameAuthor { get; set; }
        public string Title { get; set; }
        public int Date { get; set; }
        public string Placing { get; set; }
        public  int id { get; set; }
        public BooksClass(int id,string IDBook, string NameAuthor, string Title, int Date, string Placing)
        {
            this.id = id;
            this.IDBook = IDBook;
            this.NameAuthor = NameAuthor;
            this.Title = Title;
            this.Date = Date;
            this.Placing = Placing;
        }
    }
}
